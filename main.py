"""
Copyright © 2021 Chengyuan Ma

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the “Software”), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE. """

import base64
import io
import os
import tempfile
import time
from contextlib import contextmanager
from typing import Optional, Generator, Union

import click
import fitz
import pytesseract
from PIL import Image
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait
from tqdm import tqdm

StrPath = Union[str, os.PathLike]


def clean_doc_online(images: Generator[StrPath, None, None], browser: str) \
        -> Generator[Image.Image, None, None]:
    """
    Cleans the scanned document pages using docleaner's online service.

    :param images: A generator yielding paths to document pages.
    :param browser: Browser type, can be "chrome", "firefox", "safari", or
        "edge". Requires the browser and its webdriver to be installed.
    """
    if browser == "chrome":
        browser = webdriver.Chrome()
    elif browser == "firefox":
        browser = webdriver.Firefox()
    elif browser == "safari":
        browser = webdriver.Safari()
    elif browser == "edge":
        browser = webdriver.Edge()
    else:
        raise RuntimeError("Unknown browser type")
    # Timeout for web driver waits. 10s is a reasonable value unless you have
    # a very high-res image / terrible network.
    timeout = 10
    browser.get("http://www.docleaner.cn/experience.html")

    # Turn on background removal and automatic deskewing.
    WebDriverWait(browser, timeout).until(
        expected_conditions.visibility_of_element_located(
            (By.XPATH,
             "//input[@value='去背景']/parent::div/preceding-sibling::button")
        )
    ).click()
    WebDriverWait(browser, timeout).until(
        expected_conditions.visibility_of_element_located(
            (By.XPATH,
             "//input[@value='自动纠斜']/parent::div/preceding-sibling::button")
        )
    ).click()
    # Wait for a while to ensure the changes take effect.
    time.sleep(1)

    uploader = WebDriverWait(browser, timeout).until(
        expected_conditions.presence_of_element_located(
            (By.CLASS_NAME, "layui-upload-file")))

    try:
        uploader.send_keys(next(images))
        while True:
            # Write like this instead of a for loop enables us to fetch the
            # next image while the browser & remote server are processing the
            # image just uploaded. Converting a pdf page to an image is slow,
            # so we here save a lot of time :)
            next_image = next(images)
            # Wait for the result image to be visible.
            result = WebDriverWait(browser, timeout).until(
                expected_conditions.visibility_of_element_located(
                    (By.ID, "dragImgRight")))
            # Hide the result image again so the wait condition above can be
            # re-used.
            browser.execute_script(
                "arguments[0].parentNode.classList.add('layui-hide');", result)
            result = result.get_attribute("src")
            result = base64.b64decode(
                result.replace("data:image/jpg;base64,", ""))
            yield Image.open(io.BytesIO(result))
            if next_image == "":
                # See convert_pdf_to_images for the reason behind this weird
                # branch.
                break
            uploader.send_keys(next_image)
    except StopIteration:
        pass

    browser.quit()


def convert_pdf_to_images(pdf: StrPath, fmt: str, dpi: int,
                          output: Optional[StrPath] = None,
                          first_page: Optional[int] = None,
                          last_page: Optional[int] = None) \
        -> Generator[StrPath, None, None]:
    """
    Converts a pdf file to images. This a necessary pre-processing step
    because docleaner online only accepts images as inputs.

    :param pdf: The path to the pdf file.
    :param fmt: Image file format. jpg is the fastest but not lossless; png is
        lossless but slow; tiff is theoretically the best but occupies a lot of
        disk space.
    :param dpi: Pixel density of the output image.
    :param output: The output directory of intermediate images.
    :param first_page: First page to convert (starting from 1, inclusive).
    :param last_page: Last page to convert (starting from 1, inclusive).
    :return: A generator yielding the paths to the images.
    """
    doc = fitz.Document(pdf)

    @contextmanager
    def normal_dir(dir_path):
        from pathlib import Path
        Path(dir_path).mkdir(parents=True, exist_ok=True)
        yield dir_path

    matrix = fitz.Matrix(dpi / 72, dpi / 72)
    first_page = 0 if first_page is None else first_page - 1
    last_page = doc.page_count if last_page is None else last_page
    with tempfile.TemporaryDirectory() if output is None else normal_dir(
            output) as path:
        for i in range(first_page, last_page):
            filename = os.path.join(path, f"{i}.{fmt}")
            # noinspection PyUnresolvedReferences
            doc[i].get_pixmap(matrix=matrix).save(filename)
            yield filename
        if output is None:
            # Yield an empty string if we are using a temporary directory,
            # because without this, the temporary directory will be cleaned
            # up the moment the last filename is yielded, when the caller
            # hasn't done anything to the yielded temp file yet. Yielding an
            # emtpy string keeps the TemporaryDirectory object in memory
            # longer so the problem is solved.
            yield ""


def convert_images_to_pdf(images: Generator[Image.Image, None, None],
                          output: StrPath,
                          ocr: bool = True, total: Optional[int] = None):
    """
    Converts and merges images to a one-page pdf file, performing optional
    OCR in the process.

    :param images: A generator yielding PIL image objects.
    :param output: Path to the result pdf.
    :param ocr: Whether to perform OCR(Optical Character Recognition).
    :param total: An optional integer hinting the total number of images given.
        If supplied, a progress bar will be displayed during the conversion.
    """
    doc = fitz.Document()
    for image in images if total is None else tqdm(images, total=total):
        if ocr:
            pdf = pytesseract.image_to_pdf_or_hocr(image)
            page = fitz.Document(stream=pdf, filetype="pdf")
            doc.insert_pdf(page)
        else:
            # noinspection PyUnresolvedReferences
            page = doc.new_page(width=image.width, height=image.height)
            buffer = io.BytesIO()
            image.save(buffer, format="jpeg")
            page.insert_image(fitz.Rect(0, 0, image.width, image.height),
                              stream=buffer)
    doc.save(output)


# noinspection PyShadowingBuiltins
@click.command()
@click.argument("input", type=click.Path(exists=True))
@click.argument("output", type=click.Path())
@click.option("-f", "--format", default="png",
              help="Intermediate image format.")
@click.option("-d", "--dpi", default=300, help="DPI for rasterization.")
@click.option("-b", "--browser", default="chrome",
              help="The browser selenium uses.")
@click.option("--first-page", type=int, help="First page to convert/clean.")
@click.option("--last-page", type=int, help="Last page to convert/clean.")
@click.option("--ocr/--no-ocr", default=True,
              help="Whether to perform OCR during the conversion.")
@click.option("--clean/--dont-clean", default=True,
              help="Whether to clean pdf using docleaner's online service.")
def main(input: str, output: str, format: str, dpi: int, browser: str,
         first_page: Optional[int], last_page: Optional[int], ocr: bool,
         clean: bool):
    images = convert_pdf_to_images(input, fmt=format, dpi=dpi,
                                   first_page=first_page, last_page=last_page)
    if clean:
        images = clean_doc_online(images, browser)
    doc = fitz.Document(input)
    total = (doc.page_count if last_page is None else last_page) \
        - (0 if first_page is None else first_page - 1)
    convert_images_to_pdf(images, output, ocr=ocr, total=total)


if __name__ == "__main__":
    main()

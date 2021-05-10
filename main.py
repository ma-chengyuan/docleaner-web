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
from typing import Optional, Generator, Union, Tuple

import click
import fitz
import pytesseract
import requests
from PIL import Image
from tqdm import tqdm

StrPath = Union[str, os.PathLike]


def clean_doc_selenium(images: Generator[Tuple[bytes, str], None, None]) \
        -> Generator[Image.Image, None, None]:
    """
    Cleans the scanned document pages using docleaner's online service.

    :param images: A generator yielding paths to document pages.
    """
    for (image, ext) in images:
        # noinspection HttpUrlsUsage
        req = requests.post("http://service.docleaner.cn/attachCollect/upload",
                            files={"file": (f"image.{ext}", image)})
        data = {
            # Weird typo in the API.
            "paramers": "降噪,去斑点,去黑边,去背景,自动纠斜",
            "type": "image",
            "storePath": req.json()["data"]["storePath"],
            "userId": ""
        }
        # noinspection HttpUrlsUsage
        req = requests.post("http://service.docleaner.cn/exe/daqw", data=data)
        result = base64.b64decode(req.json()["data"]["outFileStr"])
        yield Image.open(io.BytesIO(result))


def convert_pdf_to_images(pdf: StrPath, fmt: str, dpi: int,
                          output: Optional[StrPath] = None,
                          first_page: Optional[int] = None,
                          last_page: Optional[int] = None) \
        -> Generator[Tuple[bytes, str], None, None]:
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
    :return: A generator yielding (image as raw bytes, image format).
    """
    doc = fitz.Document(pdf)

    matrix = fitz.Matrix(dpi / 72, dpi / 72)
    first_page = 0 if first_page is None else first_page - 1
    last_page = doc.page_count if last_page is None else last_page
    for i in range(first_page, last_page):
        # noinspection PyUnresolvedReferences
        data = doc[i].get_pixmap(matrix=matrix).tobytes()
        if output is not None:
            filename = os.path.join(output, f"{i}.{fmt}")
            with open(filename, "wb") as file:
                file.write(data)
        yield data, "png"


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
@click.option("--first-page", type=int, help="First page to convert/clean.")
@click.option("--last-page", type=int, help="Last page to convert/clean.")
@click.option("--ocr/--no-ocr", default=True,
              help="Whether to perform OCR during the conversion.")
@click.option("--clean/--dont-clean", default=True,
              help="Whether to clean pdf using docleaner's online service.")
def main(input: str, output: str, format: str, dpi: int,
         first_page: Optional[int], last_page: Optional[int], ocr: bool,
         clean: bool):
    images = convert_pdf_to_images(input, fmt=format, dpi=dpi,
                                   first_page=first_page, last_page=last_page)
    if clean:
        images = clean_doc_selenium(images)
    doc = fitz.Document(input)
    total = (doc.page_count if last_page is None else last_page) \
        - (0 if first_page is None else first_page - 1)
    convert_images_to_pdf(images, output, ocr=ocr, total=total)


if __name__ == "__main__":
    main()

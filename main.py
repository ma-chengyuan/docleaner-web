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
import glob
import io
import os
from collections.abc import Iterable
from itertools import repeat
from multiprocessing import Pool
from pathlib import Path
from typing import Optional, Union, Tuple

import click
import fitz
import pytesseract
import requests
from PIL import Image
from tqdm import tqdm

StrPath = Union[str, os.PathLike]


def merge_to_pdf(pages: Iterable[Union[Image.Image, bytes]], output: StrPath):
    """
    Converts and merges images to a one-page pdf file, performing optional
    OCR in the process.

    :param pages: A generator yielding PIL image objects.
    :param output: Path to the result pdf.
    """
    doc = fitz.Document()
    for page in pages:
        if isinstance(page, Image.Image):
            # noinspection PyUnresolvedReferences
            doc_page = doc.new_page(width=page.width, height=page.height)
            buffer = io.BytesIO()
            page.save(buffer, format="jpeg")
            doc_page.insert_image(fitz.Rect(0, 0, page.width, page.height),
                              stream=buffer)
        else:
            page = fitz.Document(stream=page, filetype="pdf")
            doc.insert_pdf(page)
    doc.save(output)


def convert_pdf_page_to_image(pdf: StrPath, idx: int, dpi: int) -> bytes:
    """
    Converts a PDF page to an image.

    :param pdf: Path to the PDF.
    :param idx: Page index (0-offset).
    :param dpi: Pixel density. A value > 200 is recommended.
    :return: Raw image as bytes.
    """
    doc = fitz.Document(pdf)
    matrix = fitz.Matrix(dpi / 72, dpi / 72)
    assert 0 <= idx < doc.page_count
    # noinspection PyUnresolvedReferences
    return doc[idx].get_pixmap(matrix=matrix).tobytes()


def clean_single_page(args: Tuple[StrPath, int, int, bool, bool]) \
        -> Union[Image.Image, bytes]:
    """
    Cleans a single page.

    :param args: A tuple consisting of (in order):
        1. Path to the page (pdf or image),
        2. Index (image index or page index in PDF),
        3. DPI (-1 if an image is direcly supplied),
        4. Whether to perform OCR,
        5. Whether to actually clean the page.
    :return:  If OCR is enabled, a OCR-ed PDF in raw bytes, otherwise an PIL
        Image object representing the cleaned page.
    """
    page, idx, dpi, ocr, clean = args
    if dpi > 0:
        image = convert_pdf_page_to_image(page, idx, dpi)
        ext = ".png"
    else:
        with open(page, "rb") as file:
            image = file.read()
        ext = os.path.splitext(page)[1]
    if clean:
        # noinspection HttpUrlsUsage
        req = requests.post("http://service.docleaner.cn/attachCollect/upload",
                            files={"file": (f"image{ext}", image)})
        data = {
            # Weird typo in the API.
            "paramers": "降噪,去斑点,去黑边,去背景,自动纠斜",
            "type": "image",
            "storePath": req.json()["data"]["storePath"],
            "userId": ""
        }
        # noinspection HttpUrlsUsage
        req = requests.post("http://service.docleaner.cn/exe/daqw", data=data)
        image = base64.b64decode(req.json()["data"]["outFileStr"])
    image = Image.open(io.BytesIO(image))
    if ocr:
        return pytesseract.image_to_pdf_or_hocr(image)
    return image


# noinspection PyShadowingBuiltins
@click.command()
@click.argument("input", type=click.Path())
@click.argument("output", type=click.Path())
@click.option("-d", "--dpi", default=300, help="DPI for rasterization.")
@click.option("--first-page", type=int, help="First page to convert/clean.")
@click.option("--last-page", type=int, help="Last page to convert/clean.")
@click.option("--ocr/--no-ocr", default=True,
              help="Whether to perform OCR during the conversion.")
@click.option("--clean/--dont-clean", default=True,
              help="Whether to clean pdf using docleaner's online service.")
def main(input: str, output: str, dpi: int,
         first_page: Optional[int], last_page: Optional[int], ocr: bool,
         clean: bool):
    if os.path.splitext(input)[1].lower() == ".pdf":
        # PDF mode
        assert os.path.exists(input)
        page_count = fitz.Document(input).page_count
        first_page = 0 if first_page is None else first_page - 1
        last_page = page_count if last_page is None else last_page
        args = zip(repeat(input), range(first_page, last_page),
                   repeat(dpi), repeat(ocr), repeat(clean))
    else:
        # Glob mode
        files = sorted(glob.glob(input, recursive=True))
        first_page = 0 if first_page is None else first_page - 1
        last_page = len(files) if last_page is None else last_page
        args = zip(files[first_page:last_page], repeat(0), repeat(-1),
                   repeat(ocr), repeat(clean))
    total = last_page - first_page
    with Pool() as p:
        results = tqdm(p.imap(clean_single_page, args), total=total)
        if os.path.splitext(output)[1].lower() == ".pdf":
            merge_to_pdf(results, output)
        elif not os.path.exists(output) or os.path.isdir(output):
            if ocr:
                raise RuntimeError("the OCR flag is useless because we are "
                                   "writing images (not PDF) to the output "
                                   "directory.")
            if not os.path.exists(output):
                Path(output).mkdir(parents=True)
            for (index, page) in enumerate(results):
                file_path = os.path.join(output, f"{index}.jpg")
                assert isinstance(page, Image.Image)
                page.save(file_path)
        else:
            raise RuntimeError("invalid output format.")


if __name__ == "__main__":
    main()

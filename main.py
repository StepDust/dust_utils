from dust_utils.file_utils.md_to_docx import MdToDocx
from dust_utils import setup_logger
import os
import json

logger = setup_logger()


def test_md_to_word():
    md_path = r"test_config\md_to_config\test.md"
    output_path = r"C:\Users\Administered\Desktop\test.docx"
    with open(r"test_config\md_to_config\style.json", "r", encoding="utf-8") as f:
        style = json.load(f)

    with open(md_path, "r", encoding="utf-8") as f:
        md_text = f.read()

    md_to_docx = MdToDocx()
    md_to_docx.convert(md_text, output_path, style["default"])


if __name__ == "__main__":
    test_md_to_word()

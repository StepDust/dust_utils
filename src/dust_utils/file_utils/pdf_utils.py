import subprocess
import os
import tempfile
import shutil
import logging


# 配置日志
logger = logging.getLogger(__name__)


class PDFUtils:

    @staticmethod
    def _get_pdftoppm(poppler_dir):
        exe = "pdftoppm.exe" if os.name == "nt" else "pdftoppm"
        path = os.path.join(poppler_dir, exe)
        if not os.path.exists(path):
            raise FileNotFoundError(f"未找到: {path}")
        return path

    @staticmethod
    def _pdf_to_images(pdf_path, output_dir, poppler_dir, dpi="200"):
        logger.debug("开始转换 PDF 为图片")
        pdftoppm = PDFUtils._get_pdftoppm(poppler_dir)

        command = [
            pdftoppm,
            pdf_path,
            os.path.join(output_dir, "page"),
            "-png",
            "-r",
            f"{dpi}",  # 分辨率，保证视觉几乎无损
        ]

        subprocess.run(command, check=True)

    @staticmethod
    def _images_to_pdf(images_dir, output_pdf):
        # 👉 懒加载 Pillow
        from PIL import Image

        logger.debug("开始转换 图片为 PDF")

        files = sorted(f for f in os.listdir(images_dir) if f.endswith(".png"))

        if not files:
            logger.error(f"路径下没有图片：{images_dir}")
            return

        image_list = []

        for f in files:
            path = os.path.join(images_dir, f)
            img = Image.open(path)
            image_list.append(img.convert("RGB"))

        image_list[0].save(
            output_pdf, save_all=True, append_images=image_list[1:], compress_level=6
        )
        logger.success(f"保存至：{output_pdf}")

    @staticmethod
    def pdf_to_noncopyable_pdf(input_pdf, output_pdf, poppler_dir, dip="200"):
        """
        转换为不可复制PDF（视觉基本无差别）
        """

        temp_dir = tempfile.mkdtemp(prefix="pdf_img_")

        try:
            # 1. PDF -> 图片
            PDFUtils._pdf_to_images(input_pdf, temp_dir, poppler_dir, dip)

            # 2. 图片 -> PDF
            PDFUtils._images_to_pdf(temp_dir, output_pdf)

        finally:
            # 清理
            shutil.rmtree(temp_dir, ignore_errors=True)

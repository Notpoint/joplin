import os
from docx2pdf import convert
import fitz  # PyMuPDF
from PIL import Image
import shutil # 用于删除文件夹

def convert_word_to_image_pdf_simple(docx_path):
    """
    使用docx2pdf库将Word文档转换为图片型PDF，代码更简洁。
    依赖: Microsoft Word (Windows) 或 LibreOffice (macOS/Linux)。

    :param docx_path: 输入的Word文档路径。
    """
    # --- 检查输入文件 ---
    if not os.path.exists(docx_path):
        print(f"错误：文件 '{docx_path}' 不存在。")
        return

    # --- 定义文件路径 ---
    file_name = os.path.splitext(os.path.basename(docx_path))[0]
    dir_path = os.path.dirname(docx_path)
    temp_pdf_path = os.path.join(dir_path, f"{file_name}_temp.pdf")
    image_pdf_path = os.path.join(dir_path, f"{file_name}_image.pdf")
    image_folder = os.path.join(dir_path, f"{file_name}_images_temp")

    try:
        # --- 步骤 1: 直接根据路径将Word转换为标准PDF ---
        print("步骤 1/4: 正在将DOCX转换为标准PDF...")
        # docx2pdf 封装了后台打开Word并保存的过程
        convert(docx_path, temp_pdf_path)
        print(f"标准PDF已生成: {temp_pdf_path}")

        # --- 步骤 2: 将PDF页面转换为图片 ---
        print("步骤 2/4: 正在将PDF页面转换为图片...")
        if not os.path.exists(image_folder):
            os.makedirs(image_folder)

        pdf_doc = fitz.open(temp_pdf_path)
        image_paths = []
        for page_num in range(len(pdf_doc)):
            page = pdf_doc.load_page(page_num)
            pix = page.get_pixmap(dpi=300)
            image_path = os.path.join(image_folder, f"page_{page_num:03d}.png")
            pix.save(image_path)
            image_paths.append(image_path)
        pdf_doc.close()
        print("页面图片已生成。")

        # --- 步骤 3: 将图片合并为图片型PDF ---
        print("步骤 3/4: 正在将图片合并为图片型PDF...")
        if not image_paths:
            raise ValueError("未生成任何图片。")

        first_image = Image.open(image_paths[0]).convert("RGB")
        other_images = [Image.open(p).convert("RGB") for p in image_paths[1:]]
        first_image.save(image_pdf_path, "PDF", resolution=100.0, save_all=True, append_images=other_images)
        print(f"图片型PDF已生成: {image_pdf_path}")

    except Exception as e:
        print(f"处理过程中发生错误: {e}")

    finally:
        # --- 步骤 4: 清理临时文件 ---
        print("步骤 4/4: 正在清理临时文件...")
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
        if os.path.exists(image_folder):
            shutil.rmtree(image_folder) # shutil.rmtree可以删除非空文件夹
        print("清理完成。")

if __name__ == '__main__':
    # 在这里输入你的Word文档路径
    input_path = "C:\\Users\\13684\\Desktop\\jophin\\笔记集合.docx"
    convert_word_to_image_pdf_simple(input_path)
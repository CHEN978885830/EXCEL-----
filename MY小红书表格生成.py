import os
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter

def create_excel_from_folder(main_folder_path, output_excel_path):
    # 创建一个新的Excel工作簿
    workbook = Workbook()
    sheet = workbook.active

    # 设置表头
    sheet["A1"] = "文件夹名"
    sheet["B1"] = "文本文件内容"
    sheet["C1"] = "图片"

    # 设置图片大小
    image_width_cm = 6.5
    image_height_cm = 10.5
    pixels_per_cm = 28.35  # 假设DPI为96

    # 获取主文件夹下的所有子文件夹
    subfolders = [f.path for f in os.scandir(main_folder_path) if f.is_dir()]

    # 遍历每个子文件夹
    for i, folder in enumerate(subfolders, start=2):
        # 获取文件夹名
        folder_name = os.path.basename(folder)
        sheet[f"A{i}"] = folder_name

        # 获取子文件夹中的所有txt文件路径
        txt_files = [f.path for f in os.scandir(folder) if f.is_file() and f.name.endswith(".txt")]

        # 遍历每个txt文件
        for j, txt_file_path in enumerate(txt_files, start=2):
            with open(txt_file_path, "r", encoding="utf-8") as file:
                txt_content = file.read()
            sheet[f"B{i + j - 1}"] = txt_content

        # 获取子文件夹中的所有jpg图片路径
        jpg_files = [f.path for f in os.scandir(folder) if f.is_file() and f.name.endswith(".jpg")]

        # 遍历每个jpg文件
        for k, jpg_file_path in enumerate(jpg_files, start=2):
            img = Image(jpg_file_path)
            img.width = image_width_cm * pixels_per_cm
            img.height = image_height_cm * pixels_per_cm
            sheet.add_image(img, f"C{i + k - 1}")

    # 保存Excel文件
    workbook.save(output_excel_path)

# 用法示例
main_folder_path = r"E:\cheneach\江湖\江湖工具箱\软件数据\XHS提取作品\单作品解析"
output_excel_path = r"C:\Users\chenliqi\Desktop\output.xlsx"
create_excel_from_folder(main_folder_path, output_excel_path)

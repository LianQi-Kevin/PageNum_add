import io
import os
from typing import Optional, Tuple, List

import PyPDF2
from reportlab.lib.pagesizes import A4, A3, A2, A1, A0, landscape
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

# 全局默认值
DEFAULT_MARGINS = (2.5, 2.5, 2.5, 3.0)
DEFAULT_FOOTER_HEIGHT = 1.75
DEFAULT_START_PAGE_NUMBER = 1


def print_welcome():
    print("欢迎使用PDF页码添加工具！")
    print("本程序将帮助您在PDF文件的指定位置添加页码。")


def get_page_size() -> Tuple[float, float]:
    page_size_input = input("请输入页面尺寸（A4, A3, A2, A1, A0），默认为A4 (回车即使用默认参数)：").upper()
    if page_size_input == "A3":
        return A3
    elif page_size_input == "A2":
        return A2
    elif page_size_input == "A1":
        return A1
    elif page_size_input == "A0":
        return A0
    else:
        return A4


def get_page_number_position() -> str:
    position_input = input(
        "请选择页码放置的位置（top, bottom, left, right, auto），默认为auto (回车即使用默认参数)：").lower()
    if position_input not in ['top', 'bottom', 'left', 'right', 'auto']:
        return 'auto'
    return position_input


def scan_pdf_files(directory: str = '.') -> List[str]:
    """扫描指定目录下的PDF文件"""
    return [f for f in os.listdir(directory) if f.lower().endswith('.pdf')]


def select_pdf_file(pdf_files: List[str], directory: str = '.') -> Optional[str]:
    """
    选择PDF文件。

    参数:
    pdf_files (List[str]): PDF文件列表。
    directory (str): 目录路径。

    返回:
    Optional[str]: 选中的PDF文件路径。
    """
    if len(pdf_files) == 1:
        print(f"检测到一个PDF文件：{pdf_files[0]}")
        confirm = input("是否使用该文件？(y/n): ")
        if confirm.lower() == 'y':
            return os.path.join(directory, pdf_files[0])
        else:
            return None

    while True:
        for idx, pdf_file in enumerate(pdf_files, 1):
            print(f"{idx}: {pdf_file}")
        print(f"{len(pdf_files) + 1}: 选择其他目录")
        choice = input("请选择一个PDF文件的编号（输入exit退出）：")
        if choice.lower() == 'exit':
            return None
        if choice.isdigit() and 1 <= int(choice) <= len(pdf_files):
            return os.path.join(directory, pdf_files[int(choice) - 1])
        elif choice == str(len(pdf_files) + 1):
            return None


def get_user_input() -> Tuple[Tuple[float, float, float, float], float, int]:
    """
    获取用户输入的页边距、页脚高度和起始页码。

    返回:
    Tuple[Tuple[float, float, float, float], float, int]: 页边距、页脚高度和起始页码。
    """
    margins_input = input(f"请输入页边距（上, 右, 下, 左），以厘米为单位，默认为{DEFAULT_MARGINS} (回车即使用默认参数)：")
    footer_height_input = input(f"请输入页脚高度，以厘米为单位，默认为{DEFAULT_FOOTER_HEIGHT} (回车即使用默认参数)：")
    start_page_number_input = input(f"请输入起始页码，默认为{DEFAULT_START_PAGE_NUMBER} (回车即使用默认参数)：")

    margins = tuple(map(float, margins_input.split(','))) if margins_input else DEFAULT_MARGINS
    footer_height = float(footer_height_input) if footer_height_input else DEFAULT_FOOTER_HEIGHT
    start_page_number = int(start_page_number_input) if start_page_number_input else DEFAULT_START_PAGE_NUMBER

    return margins, footer_height, start_page_number


def create_page_number_pdf(num_pages: int, pdf_path: str, margins: Tuple[float, float, float, float],
                           footer_height: float, start_page_number: int, page_size: Tuple[float, float],
                           position: str, font_path: Optional[str] = None, font_size: float = 10.5) -> io.BytesIO:
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=page_size)

    # 注册系统中的宋体字体
    if font_path is None:
        font_path = os.path.join(os.getenv('WINDIR'), 'FONTS', 'SIMSUN.TTC')
        assert os.path.exists(font_path), "未找到宋体文件，请检查系统字库安装情况"
    pdfmetrics.registerFont(TTFont('SimSun', font_path))

    # 打开现有的PDF文件以检测页面方向
    reader = PyPDF2.PdfReader(pdf_path)

    for i in range(num_pages):
        page = reader.pages[i]
        width, height = page.mediabox.upper_right

        if width > height:  # 横板页面
            current_page_size = landscape(page_size)
        else:  # 竖板页面
            current_page_size = page_size

        c.setPageSize(current_page_size)

        # 计算页码
        page_number = start_page_number + i

        # 根据用户选择的位置计算页码的位置
        if position == 'top':
            x_position = (current_page_size[0] - margins[3] * cm - margins[1] * cm) / 2 + margins[3] * cm
            y_position = current_page_size[1] - footer_height * cm
            text_angle = 0
        elif position == 'bottom':
            x_position = (current_page_size[0] - margins[3] * cm - margins[1] * cm) / 2 + margins[3] * cm
            y_position = footer_height * cm
            text_angle = 0
        elif position == 'left':
            x_position = footer_height * cm
            y_position = (current_page_size[1] - margins[0] * cm - margins[2] * cm) / 2 + margins[0] * cm
            text_angle = 270
        elif position == 'right':
            x_position = current_page_size[0] - footer_height * cm
            y_position = (current_page_size[1] - margins[0] * cm - margins[2] * cm) / 2 + margins[0] * cm
            text_angle = 90
        else:  # auto
            x_position = (current_page_size[0] - margins[3] * cm - margins[1] * cm) / 2 + margins[3] * cm
            y_position = footer_height * cm
            text_angle = 0

        # 手动补偿页码位置(奇怪的对不齐)
        x_position += 0.05 * cm
        y_position += 0.15 * cm

        # 计算文字宽度
        text_width = c.stringWidth(str(page_number), 'SimSun', font_size)

        # 对于顶部和底部位置，调整x坐标以使文字居中
        if position in ['top', 'bottom', 'auto']:
            x_position -= text_width / 2

        # 绘制页码
        c.setFont('SimSun', font_size)
        c.saveState()
        c.translate(x_position, y_position)
        c.rotate(text_angle)
        c.drawString(0, 0, str(page_number))
        c.restoreState()
        c.showPage()

    c.save()
    packet.seek(0)
    return packet


def add_page_numbers(pdf_path: str, output_path: str, margins: Tuple[float, float, float, float], footer_height: float,
                     start_page_number: int, page_size: Tuple[float, float], position: str) -> str:
    # 打开现有的PDF文件
    reader = PyPDF2.PdfReader(pdf_path)
    writer = PyPDF2.PdfWriter()
    num_pages = len(reader.pages)

    # 创建页码PDF文件
    packet = create_page_number_pdf(num_pages, pdf_path, margins, footer_height, start_page_number, page_size, position)
    temp_reader = PyPDF2.PdfReader(packet)

    for i in range(num_pages):
        # 获取原始PDF页面
        page = reader.pages[i]

        # 获取临时PDF页面（页码）
        overlay = temp_reader.pages[i]

        # 将页码页面合并到原始页面上
        page.merge_page(overlay)

        # 将合并后的页面添加到新的PDF写入器
        writer.add_page(page)

    # 保存修改后的PDF文件
    with open(output_path, "wb") as output_pdf:
        writer.write(output_pdf)

    return output_path


def check_output_path(output_path: str) -> str:
    """
    检查输出文件路径是否已存在，如果存在，提示用户输入新的文件名或确认是否覆盖。

    参数:
    output_path (str): 输出PDF文件的路径。

    返回:
    str: 最终确认的输出PDF文件路径。
    """
    while os.path.exists(output_path):
        print(f"文件 {output_path} 已存在。")
        choice = input("是否覆盖现有文件？(y/n): ")
        if choice.lower() == 'y':
            break
        output_path = input("请输入新的输出文件名（包括路径）：")
        if not output_path.lower().endswith('.pdf'):
            output_path += '.pdf'
    return output_path


def main():
    print_welcome()

    pdf_files = scan_pdf_files()
    directory = '.'

    while not pdf_files:
        print("当前目录下没有找到PDF文件。")
        choice = input("是否手动指定目录？(y/n): ")
        if choice.lower() == 'y':
            directory = input("请输入目录路径（输入exit退出）：")
            if directory.lower() == 'exit':
                return
            pdf_files = scan_pdf_files(directory)
        else:
            return

    pdf_file = select_pdf_file(pdf_files, directory)
    if not pdf_file:
        return

    margins, footer_height, start_page_number = get_user_input()
    page_size = get_page_size()
    position = get_page_number_position()

    output_pdf_file = check_output_path(f"{os.path.basename(pdf_file)}_output.pdf")
    add_page_numbers(pdf_file, output_pdf_file, margins, footer_height, start_page_number, page_size, position)
    print(f"已成功将页码添加到PDF文件。输出文件为：{output_pdf_file}")


if __name__ == "__main__":
    main()

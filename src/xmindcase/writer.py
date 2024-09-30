#!/usr/bin python3
# @File    : writer.py
# @Time    : 2024-09-14 16:40:44
# @Author  : Kelvin.Ye
import os
import shutil

from datetime import datetime

import openpyxl

from openpyxl.styles import Alignment
from openpyxl.styles import Border
from openpyxl.styles import Font
from openpyxl.styles import Side
from openpyxl.styles import numbers
from rich import print


def copy_excel(source, target_dir, target_name=None) -> str:
    """复制 Excel 文件"""
    # 判断是否为文件
    if not os.path.isfile(source):
        raise Exception(f'{source} 非文件')

    # 判断目标目录是否存在，不存在则新建
    if not os.path.exists(target_dir):
        os.mkdir(target_dir)

    # 存在 target_name 时修改复制后的文件名为 target_name
    file_name = target_name
    if not file_name:
        file_name = os.path.split(source)[1]
    name, ext = os.path.splitext(file_name)
    name = datetime.now().strftime(r'[%Y-%m-%d %H.%M.%S] ') + name
    file_name = f'{name}.xlsx'
    file_path = os.path.join(target_dir, file_name)

    # 判断目标文件是否存在
    if os.path.exists(file_path):
        raise Exception(f'{file_path} 文件已存在')

    # 复制文件
    shutil.copyfile(source, file_path)
    return file_path


def add_cell_borders(sheet):
    # 获取已使用的区域
    used_range = sheet.calculate_dimension()
    # 获耿边框样式
    border_style = Side(border_style='thin', color='000000')  # 细线黑色
    # 给每个单元格添加边框
    for row in sheet[used_range]:
        for cell in row:
            cell.border = Border(top=border_style, bottom=border_style, left=border_style, right=border_style)


def get_maxlen(dataset: list, colname: str):
    return max([len(data[colname].encode('utf-8')) for data in dataset])

def write_to_excel(wb: openpyxl.Workbook, excel_sheet: str, testcases: list):
    """写入excel"""
    # 复制模板页
    sheet = wb.copy_worksheet(wb['template'])
    sheet.title = excel_sheet
    # 遍历写入数据
    for rownum, case in enumerate(testcases):
        rownum = rownum + 2
        print(f'LineNo.{rownum} Testcas: {case}')
        # 模块
        sheet[f'B{rownum}'].value = case['module']
        sheet[f'B{rownum}'].alignment = Alignment(vertical='center', wrapText=True)
        # 路径
        sheet[f'C{rownum}'].value = case['path']
        sheet[f'C{rownum}'].alignment = Alignment(vertical='center', wrapText=True)
        # 分类
        sheet[f'D{rownum}'].value = case['cate']
        sheet[f'D{rownum}'].alignment = Alignment(vertical='center', wrapText=True)
        # 用例名称
        sheet[f'E{rownum}'].value = case['title']
        sheet[f'E{rownum}'].alignment = Alignment(vertical='center', wrapText=True)
        # 前置条件
        sheet[f'F{rownum}'].value = case['pre']
        sheet[f'F{rownum}'].alignment = Alignment(vertical='center', wrapText=True)
        # 测试数据
        sheet[f'G{rownum}'].value = case['data']
        sheet[f'G{rownum}'].alignment = Alignment(vertical='center', wrapText=True)
        # 用例步骤
        sheet[f'H{rownum}'].value = case['step']
        sheet[f'H{rownum}'].alignment = Alignment(vertical='center', wrapText=True)
        # 预期结果
        sheet[f'I{rownum}'].value = case['exp']
        sheet[f'I{rownum}'].alignment = Alignment(vertical='center', wrapText=True)
    # 自动调整列宽（模块列和路径列）
    sheet.column_dimensions['B'].width = get_maxlen(testcases, 'module')
    sheet.column_dimensions['C'].width = get_maxlen(testcases, 'path')
    # 添加表格边框
    add_cell_borders(sheet)


def add_dashboard_to_excel(wb: openpyxl.Workbook):
    dashboard_sheet = wb['dashboard']
    result_column = 'J:J'

    sheets = [name for name in wb.sheetnames if name not in ['dashboard', 'template']]

    for rownum, sheet_name in enumerate(sheets):
        # 行号
        rownum = rownum + 3

        # 案例名称
        dashboard_sheet[f'A{rownum}'].value = sheet_name
        dashboard_sheet[f'A{rownum}'].font = Font(bold=True, size=14)
        # 总编写用例数
        dashboard_sheet[f'B{rownum}'].value = f'=IFERROR(COUNTIF(INDIRECT("\'{sheet_name}\'!E:E"), "*") - 1, 0)'
        # 需执行用例数
        dashboard_sheet[f'C{rownum}'].value = f'=IFERROR(B{rownum} - G{rownum}, 0)'

        # 通过
        dashboard_sheet[f'D{rownum}'].value = f'=COUNTIF(INDIRECT("\'{sheet_name}\'!{result_column}"), "通过")'
        # 失败
        dashboard_sheet[f'E{rownum}'].value = f'=COUNTIF(INDIRECT("\'{sheet_name}\'!{result_column}"), "失败")'
        # 阻塞
        dashboard_sheet[f'F{rownum}'].value = f'=COUNTIF(INDIRECT("\'{sheet_name}\'!{result_column}"), "阻塞")'
        # 不适用
        dashboard_sheet[f'G{rownum}'].value = f'=COUNTIF(INDIRECT("\'{sheet_name}\'!{result_column}"), "不适用")'

        # 未执行
        dashboard_sheet[f'H{rownum}'].value = f'=IFERROR(C{rownum} - (D{rownum} + E{rownum}), 0)'
        # 总完成率
        dashboard_sheet[f'I{rownum}'].value = f'=IFERROR((D{rownum} + E{rownum}) / C{rownum}, 0)'
        dashboard_sheet[f'I{rownum}'].number_format = numbers.FORMAT_PERCENTAGE
        # 总通过率
        dashboard_sheet[f'J{rownum}'].value = f'=IFERROR(D{rownum} / C{rownum}, 0)'
        dashboard_sheet[f'J{rownum}'].number_format = numbers.FORMAT_PERCENTAGE

    # 总计
    last_rownum = dashboard_sheet.max_row
    total_rownum = last_rownum + 1
    # 案例名称
    dashboard_sheet[f'A{total_rownum}'].value = '总计'
    dashboard_sheet[f'A{total_rownum}'].font = Font(bold=True, size=14)
    # 总编写用例数
    dashboard_sheet[f'B{total_rownum}'].value = f'=SUM(B3:B{total_rownum - 1})'
    # 需执行用例数
    dashboard_sheet[f'C{total_rownum}'].value = f'=SUM(C3:C{total_rownum - 1})'
    # 通过
    dashboard_sheet[f'D{total_rownum}'].value = f'=SUM(D3:D{total_rownum - 1})'
    # 失败
    dashboard_sheet[f'E{total_rownum}'].value = f'=SUM(E3:E{total_rownum - 1})'
    # 阻塞
    dashboard_sheet[f'F{total_rownum}'].value = f'=SUM(F3:F{total_rownum - 1})'
    # 不适用
    dashboard_sheet[f'G{total_rownum}'].value = f'=SUM(G3:G{total_rownum - 1})'
    # 未执行
    dashboard_sheet[f'H{total_rownum}'].value = f'=SUM(H3:H{total_rownum - 1})'
    # 总完成率
    dashboard_sheet[f'I{total_rownum}'].value = f'=IFERROR((D{total_rownum} + E{total_rownum}) / C{total_rownum}, 0)'
    dashboard_sheet[f'I{total_rownum}'].number_format = numbers.FORMAT_PERCENTAGE
    # 总通过率
    dashboard_sheet[f'J{total_rownum}'].value = f'=IFERROR(D{total_rownum} / C{total_rownum}, 0)'
    dashboard_sheet[f'J{total_rownum}'].number_format = numbers.FORMAT_PERCENTAGE
    # 添加边框
    add_cell_borders(dashboard_sheet)

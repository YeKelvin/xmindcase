#!/usr/bin python3
# @File    : main.py
# @Time    : 2024-09-14 16:40:37
# @Author  : Kelvin.Ye
import os
import sys

import openpyxl
import typer

from rich import print

from xmindcase.config import template_excel_path
from xmindcase.parser import node_to_case
from xmindcase.parser import parse_xmind
from xmindcase.parser import validate_nodes
from xmindcase.writer import add_dashboard_to_excel
from xmindcase.writer import copy_excel
from xmindcase.writer import write_to_excel


# 添加项目路径到 system-path
sys.path.append(os.path.dirname(sys.path[0]))


def xmind_to_excel(xmind_file_path: str, xmind_sheet_name: str|None=None, debug=False):
    """XMind 转 Excel"""
    # 获取 xmind 文件名
    file_name = os.path.split(xmind_file_path)[1]
    xmind_name, file_ext = os.path.splitext(file_name)
    # 解析 XMind
    print('加载 xmind 文件')
    sheets = parse_xmind(xmind_file_path)
    # 仅转换指定的 sheet 页
    if xmind_sheet_name:
        sheets = [sheet for sheet in sheets if sheet['title'] == xmind_sheet_name]
        if not sheets:
            raise Exception('指定的 sheet 页不存在')
    suites = []
    # 遍历 sheet 页转换
    for sheet in sheets:
        print(f'开始转换 sheet 页: {sheet["title"]}')
        # 根节点
        root = sheet['topic']
        # 根名称
        name = root['title']
        # 子节点
        nodes = root['topics']
        # 用例列表
        cases = []
        # 用例元数据
        metadata = {
            'root': name,   # 根节点
            'module': [],   # 模块
            'path': [],     # 路径
            'cate': [],     # 分类
            'title': [],    # 用例标题
            'pre': [],      # 前置条件
            'data': [],     # 测试数据
            'step': [],     # 测试步骤
            'exp': [],      # 预期结果
            'note': []      # 备注
        }
        # 校验节点规范
        validate_nodes(nodes)
        # 节点数据转为用例数据
        node_to_case(nodes, cases, metadata)
        # 添加至用例集
        suites.append({'sheet': sheet['title'], 'cases': cases})
        if debug:
            [print(case) for case in cases]
    # XMindCase 解析完成
    print(f'xmindcase 解析完成，共 {sum([len(suite['cases']) for suite in suites])} 条用例\n')
    # 复制测试用例模板文件
    output_path = copy_excel(template_excel_path, target_name=f'{xmind_name}.xlsx')
    print('写入 excel 开始\n')
    # 打开 excel
    wb = openpyxl.load_workbook(output_path)
    # 写入 excel
    for suite in suites:
        write_to_excel(wb, suite['sheet'], suite['cases'])
    add_dashboard_to_excel(wb)
    # 删除 sheet 模板
    wb.remove(wb['template'])
    # 保存 excel
    wb.save(output_path)
    # 关闭 excel
    wb.close()
    print('\n写入 excel 完成\n')
    print(f'测试用例输出路径: {output_path}\n')


def main(file: str, sheet: str|None = None, debug: bool = False):
    if not file:
        raise typer.BadParameter('请输入正确的 xmind 文件路径')
    xmind_to_excel(file, sheet, debug)


if __name__ == '__main__':
    typer.run(main)

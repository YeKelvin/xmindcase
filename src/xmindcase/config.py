#!/usr/bin python3
# @File    : config.py
# @Time    : 2024-09-14 16:40:53
# @Author  : Kelvin.Ye
import os


# 项目路径
project_path = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir, os.pardir))

# excel模板文件路径
template_excel_path = os.path.join(project_path, 'templates', 'testcase.template.xlsx')

# xmind模板文件路径
template_xmind_path = os.path.join(project_path, 'templates', 'testcase.example.xmind')

# 默认输出路径
default_output_path = os.path.join(project_path, 'outputs')

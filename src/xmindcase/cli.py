import os
import sys

import typer

from xmindcase.converter import xmind_to_excel


# 添加项目路径到 system-path
sys.path.append(os.path.dirname(sys.path[0]))


def main(file: str, sheet: str|None = None, output: str|None = None, debug: bool = False):
    if not file:
        raise typer.BadParameter('请输入正确的 xmind 文件路径')
    xmind_to_excel(file, sheet, output, debug)


def run():
    typer.run(main)


if __name__ == '__main__':
    run()

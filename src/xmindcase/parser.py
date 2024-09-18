#!/usr/bin python3
# @File    : parser.py
# @Time    : 2024-09-14 16:40:21
# @Author  : Kelvin.Ye
from xmindparser import xmind_to_dict


# 标签
TAGS = ['module', 'path', 'cate', 'title', 'pre', 'data', 'step', 'exp', 'note']


def parse_xmind(file_path: str) -> list[dict]:
    """解析 xmind 文件"""
    sheets = xmind_to_dict(file_path)
    return sheets


def validate_nodes(nodes: list):
    """校验节点是否符合规范

    exp 和 note 不允许有子节点

    Args:
        nodes (list): 节点列表
    """
    # 方法名改为校验topic格式
    for node in nodes:
        # 中文冒号替换为英文冒号
        title = node['title'].replace('：', ':')
        if ':' not in title and (
            title.startswith('module') or  # noqa
            title.startswith('path') or  # noqa
            title.startswith('cate') or  # noqa
            title.startswith('title') or  # noqa
            title.startswith('pre') or  # noqa
            title.startswith('data') or  # noqa
            title.startswith('step') or  # noqa
            title.startswith('exp') or  # noqa
            title.startswith('note')  # noqa
        ):
            raise Exception(f'topic:[ {title} ] 格式不正确')
        if 'topics' in node:
            validate_nodes(node['topics'])


def parse_node(node: dict):
    """解析节点

    Args:
        node (dict): 节点
    """
    # 中文冒号替换为英文冒号
    data = node['title'].replace('：', ':')
    # 分割标签和内容
    splits = data.split(':')
    hastag = False
    tag = ''
    text = ''
    if len(splits) >= 2:
        tag = splits[0]
        tag = tag.strip()  # 移除首尾空格
        if tag in TAGS:
            hastag = True
            text = ':'.join(splits[1:])
            text = text.strip()  # 移除首尾空格
    return hastag, tag, text


def node_to_case(nodes: list, cases: list, metadata: dict) -> None:
    """节点转用例

    Args:
        nodes (list): 节点列表
        cases (list): 用例列表
        metadata (dict): 元数据
    """
    for node in nodes:
        # 解析主题
        hastag, tag, text = parse_node(node)
        # 添加用例原始数据
        if hastag:
            metadata[tag].append(text)
        # 存在子节点时，递归解析
        if 'topics' in node:
            node_to_case(node['topics'], cases, metadata)

        # 没有"topics"节点则代表已递归至路径末端，开始组装数据并添加至用例集
        # 路径上存在 title 才识别为一条用例
        if metadata['title'] and metadata['exp']:
            module = '-'.join(metadata['module'])
            path = '-'.join(metadata['path'])
            cate = '-'.join(metadata['cate'])
            title = '-'.join(metadata['title'])
            code = hash(f'{module}:{path}:{cate}:{title}')

            # 抵达路径末端时，判断用例是否已存在，存在则拼接预期结果，不存在则添加用例
            # path、cat 和 title 相同代表末端有多个 exp （预期结果）
            match = [case for case in cases if case['code'] == code]
            if match:
                existed_case = match[0]
                existed_case['exp'] = existed_case['exp'] + '\n' + '-'.join(metadata['exp'])
                existed_case['note'] = existed_case['note'] + '\n' + '-'.join(metadata['note'])
            else:
                cases.append({
                    'code': code,
                    'module': module,
                    'path': path,
                    'cate': cate,
                    'title': title,
                    'pre': '-'.join(metadata['pre']),
                    'data': '-'.join(metadata['data']),
                    'step': '-'.join(metadata['step']),
                    'exp': '-'.join(metadata['exp']),
                    'note': '-'.join(metadata['note'])
                })
        # 回溯时删除数据
        if hastag:
            metadata[tag].pop()

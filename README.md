# XMindCase

## 安装

```bash
# 默认用 github
pip install git+https://github.com/yekelvin/xmindcase.git
# github 太慢就用 gitee 吧
pip install git+https://gitee.com/kelvinye/xmindcase.git
```

## 开发

```bash
git clone https://github.com/yekelvin/xmindcase.git
or
git clone https://gitee.com/kelvinye/xmindcase.git
cd xmindcase
# 不用 rye 用 pyenv 也行，当然直接 pip 也可以
rye sync
or
pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple
```

## xmind 用例格式说明

详情请参考

- [Github](https://github.com/yekelvin/xmindcase/blob/master/templates/testcase.example.xmind)
- [Gitee](https://gitee.com/kelvinye/xmindcase/blob/master/templates/testcase.template.xmind)

## 命令行使用说明

```bash
xmindcase excel '文件路径' [--sheet 'sheet名称'] [--output '输出路径'] [--debug] [--help]

e.g.: xmindcase excel '/xxx/testcase.xmind'
```

## 代码内使用说明

```python
# 入口在 src/xmindcase/main.py 文件的 xmind_to_excel 函数
```

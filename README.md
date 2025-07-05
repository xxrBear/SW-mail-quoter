

# 申万宏源证券 - Excel 处理与报价邮件回复

![Static Badge](https://img.shields.io/badge/build-python3.8%2B-blue?style=flat&logo=python)
![Static Badge](https://img.shields.io/badge/Excel2013%2B-%231E8449?style=flat&logo=excel)

> **注意：本工具仅支持 Windows 操作系统**


## 🚀 快速开始

**安装依赖**

使用国内源安装

```bash
pip install -r requirements.txt -i https://mirrors.tuna.tsinghua.edu.cn/pypi/web/simple
```

或者使用 [uv](https://github.com/astral-sh/uv)（推荐）

```bash
uv sync
```

如果你没有安装，可以使用以下脚本

```bash
scripts\install_uv.bat
```

**开发者工具（可选）**


<details>
<summary>点击展开</summary>


项目提供了一个命令行工具`quoter`，如果你使用`uv sync`来同步依赖，那么已自动安装，如果你使用`pip`来安装依赖，请执行：
```shell
pip install -e .
```

- 如何使用
```shell
quoter --help # 调出命令行信息
```
</details>

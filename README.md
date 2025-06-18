# SW Excel 处理与报价邮件回复

> ⚠️ **注意：本工具仅支持 Windows 操作系统**


## 🚀 快速开始

### 手动初始化

**1. 安装 python 依赖**

使用国内源安装

```bash
pip install -r requirements.txt -i https://mirrors.tuna.tsinghua.edu.cn/pypi/web/simple
```

或者使用 [uv](https://github.com/astral-sh/uv)（推荐）

```bash
uv sync
```

**2. 配置环境变量**

创建 `.env` 文件

```bash
Set-Content -Path ".env" -Value ""
```

编辑 `.env`，填写你的邮箱配置

```env
EMAIL_SMTP_SERVER='你的邮箱服务器'
EMAIL_USER_NAME='你的邮箱账号'
EMAIL_USER_PASS='你的邮箱密码'
```


### 自动初始化

**1. 初始化依赖**

```bat
scripts\init.bat
```

**2. 配置.env文件**
```
EMAIL_SMTP_SERVER='你的邮箱服务器'
EMAIL_USER_NAME='你的邮箱账号'
EMAIL_USER_PASS='你的邮箱密码'
```

**3. 运行脚本**

```bat
scripts\run.bat
```

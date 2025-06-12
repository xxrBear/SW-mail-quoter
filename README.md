# sw-Email

## 快速开始

### 安装依赖
```shell
pip install -r requirements.txt -i https://mirrors.tuna.tsinghua.edu.cn/pypi/web/simple
```

- 或者使用 uv
```
uv sync
```

### 配置环境变量
```
# 创建 .env 文件
touch .env

# 配置 MAIL 账户密码
EMAIL_SMTP_SERVER='邮箱服务器'
EMAIL_USER_NAME='你的邮箱'
EMAIL_USER_PASS='你的密码'
```

# 🎓 理工学堂助手

**理工学堂助手** 是一个为理工学堂平台开发的小工具，帮助用户管理课程、查看作业、收集题目并一键提交成绩。

## 🔗 教程网页

### https://lgxt.zygame1314.site

## 📥 下载与使用

### 1. 直接运行 exe 文件

如果你不想安装 Python 和依赖，可以直接运行打包好的 `理工学堂助手.exe` 文件：

- 下载并运行 `理工学堂助手.exe` 文件，双击即可打开使用。

### 2. 使用油猴脚本

[脚本猫](https://scriptcat.org/zh-CN/script-show-page/2774)

### 3. 使用 Python 运行源代码

如果你希望查看或修改源码，或者直接运行 Python 文件，可以按以下步骤操作：

#### 环境要求

- 安装 Python 3.7 或更高版本。

#### 运行步骤

1. **克隆仓库或下载源代码**：

   ```bash
   git clone https://github.com/zygame1314/LGXT-Assistant.git
   cd LGXT-Assistant
   ```

2. **安装依赖**：

   依赖是程序所需的库，使用以下命令安装：

   ```bash
   pip install -r requirements.txt
   ```

3. **运行程序**：

   ```bash
   python 理工学堂助手.py
   ```

## 🛠️ 功能介绍

- **登录**：输入理工学堂用户名和密码，支持记住密码功能。
- **查看课程和作业**：快速查看当前课程和作业列表。
- **收集作业题目**：自动收集作业题目并保存为 Word 文档，包含题目图片和答案。
- **一键提交100分**：轻松将作业成绩设置为100分并提交。
- **一键提交全部课程100分**：将所有课程的作业成绩批量设置为100分并提交。

## 依赖库

如果你选择使用 Python 运行程序，请确保安装以下依赖库：

```txt
configparser - 用于处理配置文件
io - 用于处理二进制数据流
os - 用于与操作系统交互
base64 - 用于解码嵌入的图像数据
threading - 用于多线程操作
webbrowser - 用于打开网页链接
tkinter - GUI基础库
keyring - 用于安全存储密码
requests - 用于HTTP请求
concurrent.futures - 用于并行处理(ThreadPoolExecutor, as_completed)
time - 用于时间相关操作
tkinter.font - 用于字体处理
PIL (Pillow) - 图像处理库
docx (python-docx) - Word文档处理
reportlab - PDF生成库
ttkbootstrap - 美化tkinter界面
```

可以通过以下命令自动安装所有依赖：

```bash
pip install -r requirements.txt
或
pip install configparser pillow python-docx reportlab requests keyring ttkbootstrap
```

## ⚠️ 注意事项

- 该工具仅用于学习和交流，请勿用于商业用途。
- 请妥善保管你的登录信息，确保账号和密码安全。

- ## 许可证

本项目基于 [MIT License](./LICENSE) 许可发布。详细信息请参阅 LICENSE 文件。


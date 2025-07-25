# 洛谷刷题日记生成器 🚀

![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Open Source](https://img.shields.io/badge/Open%20Source-%E2%9D%A4%EF%B8%8F-brightgreen)

> "记录刷题轨迹，见证算法成长！" - 让洛谷训练日记不再枯燥乏味

## 项目简介

**洛谷刷题日记生成器**是一款专为信息学奥赛选手设计的自动化工具，它能自动从洛谷平台抓取你的提交记录，生成精美格式的刷题日记表格。告别手动记录的烦恼，让训练过程一目了然！

## ✨ 特色功能

- 🚀 **一键获取**：只需输入Cookie信息，自动获取洛谷提交记录
- 📊 **双格式输出**：同时生成美观的Excel表格和简洁的CSV文件
- 🕒 **智能排序**：提交记录按时间升序排列（越早越靠上）
- 🎨 **自动美化**：Excel表格自动着色（AC绿/WA红/TLE橙等）
- 📝 **预留笔记**：预留解题时间和笔记列，方便手动补充思考过程
- ⏱️ **时间戳命名**：自动生成带时间戳的文件名，避免覆盖历史记录

## 许可证选择

本项目采用 **MIT 许可证**，这意味着你可以：

- ✅ 自由使用、复制和修改代码
- ✅ 用于个人或商业项目
- ✅ 分发修改后的版本
- ✅ 唯一要求是保留原始许可证和版权声明

```text
MIT License

Copyright (c) 2025 [TTHILLTT]

Permission is hereby granted...
```

## 快速开始

### 安装依赖

```bash
pip install requests openpyxl
```

### 获取Cookie信息
1. 登录[洛谷](https://www.luogu.com.cn)
2. 按F12打开开发者工具
3. 转到"网络"(Network)标签页
4. 刷新页面，复制任意请求中的：
   - `__client_id`
   - `_uid`
(如果你用过VJudge，你应该知道怎么做)
## 联系作者

有疑问或建议？欢迎联系：
📧 Email: kaoxiqi@qq.com
💬 WeChat: kaoxiqi

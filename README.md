# Learning Claude Code — Claude Code 安装与使用指南

> Claude Code 命令行工具的完整安装教程、命令指南和配套资源。

---

## 📂 目录结构

```
learning-claude-code/
├── images/
│   ├── image.png ~ image-17.png     # 安装教程配套截图（共 17 张）
├── Claude_Code_安装教程.md           # 安装教程（图文版）
├── Claude_Code_安装教程.pptx          # 安装教程 PPT（深色主题，专业排版）
├── Claude Code 命令指南.md           # 日常使用命令速查
├── generate_ppt.py                  # PPT 生成脚本（python-pptx）
└── README.md                        # 本文件
```

---

## 🚀 快速开始

1. **安装 Claude Code** → [`Claude_Code_安装教程.md`](./Claude_Code_安装教程.md)
   - Node.js 安装
   - Claude Code CLI 安装
   - CC-Switch 切换国产大模型 API（DeepSeek）
   - VS Code 集成配置
2. **日常命令查阅** → [`Claude Code 命令指南.md`](./Claude%20Code%20命令指南.md)
3. **演示分享** → [`Claude_Code_安装教程.pptx`](./Claude_Code_安装教程.pptx)

---

## 📖 内容概览

### Claude_Code_安装教程.md

手把手安装教程，包含详细截图：

1. **安装 Node.js** — Claude Code 的运行环境
2. **安装 Claude Code CLI** — `npm install -g @anthropic-ai/claude-code`
3. **安装 Git** — 版本控制依赖
4. **配置 CC-Switch** — 一键切换 DeepSeek 等国产大模型 API
5. **VS Code 集成** — 在编辑器中可视化使用 Claude Code
6. **常见问题排查** — PowerShell 执行策略、DeepSeek API 认证等

### Claude Code 命令指南.md

从入门到进阶的完整命令参考：

- 界面布局认识
- 基础操作（启动、引用文件、执行命令）
- 斜杠命令速查（`/help`、`/clear`、`/compact` 等）
- 权限模式说明
- 常用工作流和技巧

### generate_ppt.py

使用 `python-pptx` 库自动生成安装教程 PPT：
- 深色主题专业排版
- 自动嵌入截图
- 支持自定义颜色方案

---

## 🛠 环境要求

- Node.js 18+
- Git 2.x
- Python 3.x（仅 PPT 生成脚本需要）
- `pip install python-pptx Pillow`（仅 PPT 生成脚本需要）

---

## 📝 许可

个人学习资料，欢迎分享和参考。

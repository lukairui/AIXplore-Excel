<div align="center">

# 📊 AIXplore-Excel

**AI 数据分析工作台 · 桌面版**

上传 Excel → 自然语言生成计算列 / 分析视图 / 对话式分析报告 → 导出增强版 Excel

[![AIXplore](https://img.shields.io/badge/org-AIXplore-blue?style=flat-square)](https://github.com/lukairui/AIXplore)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow?style=flat-square)](LICENSE)
[![Electron](https://img.shields.io/badge/Electron-⚡-9feaf9?style=flat-square&logo=electron&logoColor=white)](https://www.electronjs.org/)
[![Vue 3](https://img.shields.io/badge/Vue-3-42b883?style=flat-square&logo=vuedotjs&logoColor=white)](https://vuejs.org/)

</div>

---

## ✨ 功能

- 🧮 **AI 计算列** — 一句自然语言描述规则，自动生成整列计算逻辑（含 3 轮自动纠错循环）
- 📈 **AI 分析视图** — 模型生成 Vue 3 + ECharts 交互组件，12 种预设图表类型 + 自由应用（看板 / 日历 / 甘特图 / 游戏）
- 💬 **对话分析模式** — 多文件上传（xlsx / csv / pdf / 图片）→ 前端 LangGraph 状态机自动规划（inspect → relate → planner → execute → analyze → validate → synthesize）→ 分析报告 + 图表 + 处理后的工作簿
- 📊 **原生 Excel 图表导出** — 手写 OOXML chart XML，导出的 `.xlsx` 自带**可编辑原生图表** / 仪表盘 Sheet
- 🔒 **数据不出本机** — Excel 解析、计算列执行、图表渲染、导出 100% 本地完成，项目数据存本地 IndexedDB；只有提示词发给你自己配置的模型

## 🖥 架构

```
┌────────────────────────────────────────────────┐
│ 渲染进程（Vue 3 + Element Plus）                 │
│   excel/ 工作台 UI · IndexedDB 项目持久化         │
│   LangGraph 对话分析 · 生成视图运行时（Babel 编译） │
├────────────── IPC Bridge ──────────────────────┤
│ 主进程（Node.js）                                │
│   LLM 代理：OpenAI 兼容 API 直连                  │
│   · /chat/completions 与 /responses 双协议       │
│   · 多 API Key 随机轮换 · 3 次重试退避            │
│   · SSE 流式转发 · Qwen 系自动禁用思考            │
│   模型配置：userData/config.json（本机存储）        │
└────────────────────────────────────────────────┘
```

自带 Key：在「模型设置」里添加任意 OpenAI 兼容服务（Base URL / API Key / 模型名），支持多个模型与多 Key 轮换。

## 🚀 快速开始

```bash
# 环境要求：Node.js ≥ 20
npm install
npm run dev        # 开发模式
npm run test       # 运行单测
npm run build      # 构建
npm run build:mac  # 打包 macOS（另有 build:win / build:linux）
```

首次启动后，点击右上角「模型设置」添加至少一个模型即可开始使用。

## 📦 打包产物

`npm run build:mac|win|linux` 输出到 `release/`。

## 🧭 相关项目

- **[AIXplore](https://github.com/lukairui/AIXplore)** — 项目目录页
- **[AIXplore-PPT](https://github.com/lukairui/AIXplore-PPT)** — AI 演示文稿生成器

## 📄 协议

[MIT](LICENSE) © 2026 lukairui

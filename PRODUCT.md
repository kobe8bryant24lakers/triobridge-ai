# TrioBridge AI · 三桥 AI

<!-- impeccable:product-schema 1 -->

## Platform

web

## Stack

本次实现采用 TypeScript、React、Vite 和 Office.js。此技术选择由实现者根据 README 的三应用需求确定。

## Product Purpose

让 Word、Excel、PowerPoint 连接用户自己的 OpenAI 兼容模型。用户读取选区、选择任务、查看生成结果，然后确认写回。

## Operating Context

用户确认同时支持 Mac 和 Windows，以及无需 Microsoft 365 订阅的 Microsoft Office 永久授权版本。采用 XML 任务窗格插件清单、运行时能力检测和基础功能回退。同时提供不依赖 Office 的浏览器演示。实际 Office 宿主未验证前，不宣称已通过兼容性测试。

## Capabilities and Constraints

- Word：改写、摘要、翻译、生成内容与分节报告，报告支持原生表格插入；复杂对象选区不直接用纯文本替换。
- Excel：公式生成、解释、表格问答、分析建议与新建含公式的数据表；新表写入不覆盖原表，保留原始值与显示值的区别。
- PowerPoint：选择本地素材目录，生成 2–20 页可编辑 PPTX；读取现有页面后预览并选择文字、样式、位置与页序改动；根据表格数据插入原生图表。原有文案与大纲工具保留。
- 素材包括 DOCX、XLSX、PPTX、可提取文字的 PDF、文本、CSV/TSV 和图片。只读取用户选择的文件，在本机解析；生成时发送所选内容，图片理解另行开启。超过限制时明确报错，不静默截断。
- 多模型配置、连接测试、超时和取消；凭据不进入仓库。
- 文档读取和写入都由用户明确触发；生成本身不修改文档。
- 演示结果明确标识，不能冒充真实模型生成。
- 三个应用共用当前系统账号的多模型设置和当前选择。密钥存入系统凭据库，页面只接收元数据。并发修改会检查配置版本。
- 一个桌面版本按 Windows、Mac 和芯片架构分发。工作台从安装包中的文件打开。下载页提供安装与使用说明。
- 输入变更保留并标记旧预览，失败或取消保留既有结果；工作稿仅在内存中保留。新增功能必须区分自动化验证与真实 Office 验收，不以浏览器演示替代宿主测试。

## Brand Commitments

项目名称为 TrioBridge AI / 三桥 AI，主要界面使用中文。

## Evidence on Hand

README.md 提供产品范围；仓库尚无已有 UI、客户数据或品牌资产。演示样例均为合成内容。

## Product Principles

让文档上下文、发送内容和写入位置保持可见。用户主动生成时向所选 API 发送内容，确认后写回文档。密钥保存在系统凭据库。

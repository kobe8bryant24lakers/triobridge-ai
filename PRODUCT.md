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

- Word：改写、摘要、翻译和生成内容。
- Excel：公式生成、解释、表格问答与分析建议。
- PowerPoint：演示大纲、文案改写与内容整理。
- 多模型配置、连接测试、超时和取消；凭据不进入仓库。
- 文档读取和写入都由用户明确触发；生成本身不修改文档。
- 演示结果明确标识，不能冒充真实模型生成。
- 模型配置在各 Office WebView 中可能被隔离，提供无密钥导入导出。

## Brand Commitments

项目名称为 TrioBridge AI / 三桥 AI，主要界面使用中文。

## Evidence on Hand

README.md 提供产品范围；仓库尚无已有 UI、客户数据或品牌资产。演示样例均为合成内容。

## Product Principles

让文档上下文、发给模型的内容、目标位置和写回操作保持可见。默认仅在当前会话保留密钥，不自动发送文档，不保存生成历史。

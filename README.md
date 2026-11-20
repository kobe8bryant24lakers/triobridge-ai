# TrioBridge AI · 三桥 AI

让 Word、Excel、PowerPoint 连接你自己的 AI 模型。

TrioBridge 的 Trio 代表 Office 三件套，Bridge 代表办公应用与用户自选模型之间的桥梁。

## 项目状态

目前处于初始化阶段，已建立 Git 仓库与基础项目说明。下列能力为开发规划，尚未实现；技术栈、插件打包方式和 Office 平台兼容范围将在开发前确定。

## 产品方向

| 应用 | 首批规划能力 |
| --- | --- |
| Word | 选中文本改写、摘要、翻译与内容生成 |
| Excel | 公式生成与解释、表格问答、数据分析建议 |
| PowerPoint | 演示大纲生成、页面文案改写、内容整理 |

## 自定义模型规划

- 配置模型服务名称、API Base URL、模型 ID 和 API Key。
- 优先考虑 OpenAI 兼容接口，预留其他服务协议的适配能力。
- 保存多个模型配置，并在三件套中复用、切换。
- 提供连接测试、超时设置和清晰的错误提示。
- 规划接入用户自建或本地部署的模型服务；具体网络和平台限制待验证。
- API Key、真实文档内容和本地私有配置不提交到代码仓库。

## 初期里程碑

- [ ] 确定 Office 插件技术方案与目标平台。
- [ ] 完成模型配置、连接测试与统一调用接口。
- [ ] 打通 Word 中选中文本的读取、生成预览和确认写回。
- [ ] 扩展 Excel 与 PowerPoint 的首批能力。
- [ ] 补齐错误处理、安装说明与必要测试。

## 本地开发

在 IntelliJ IDEA 中打开克隆后的 `triobridge-ai` 目录即可管理项目。
当前尚未添加应用源码、依赖或启动脚本；确定技术方案后再补充运行步骤。

仓库基础文件：

- `README.md`：项目定位与初期规划。
- `.gitignore`：排除 IDE 文件、构建产物和本地敏感配置。
- `.editorconfig`：统一编码、换行和基础缩进。
- `LICENSE`：MIT 开源许可证。

GitHub：<https://github.com/kobe8bryant24lakers/triobridge-ai>

## License

本项目采用 [MIT License](LICENSE)。

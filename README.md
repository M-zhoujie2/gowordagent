# 智能校对

智能校对 是一个基于 **VSTO (Visual Studio Tools for Office)** 开发的 Word 插件，在原项目GOWordAgentAddIn基础上开发完成，提供 AI 智能校对功能。它将文档内容发送给 AI 模型，获取修改建议，并以修订和批注的形式写回 Word 文档。

## 功能特性

### 1. AI 问答对话
- 在侧边栏与 AI 进行实时对话
- 支持多种 AI 提供商：DeepSeek、GLM、Ollama
- 配置自动保存，支持自动连接

### 2. 智能校对
- **精准校验**：针对错别字、语病、术语不一致等精确纠错
- **全文校验**：系统性校对，包括语法、标点、用词、逻辑连贯性
- 自定义提示词，支持两种校验模式独立配置

### 3. 修订与批注
- 使用 Word 原生修订功能（TrackRevisions）
- 标记原文（删除线）和建议文本（下划线）
- 添加批注说明错误类型和修改理由
- 通过 Word 审阅面板接受或拒绝修订

### 4. 问题定位
- 聊天框显示发现的每个问题
- 点击"定位"按钮跳转到文档对应位置
- 支持严重程度标记（high/medium/low）

## 技术架构

### 核心组件

| 文件 | 行数 | 说明 |
|------|------|------|
| `GOWordAgentPaneWpf.xaml.cs` | 687 | UI 主控类，处理交互和 Word 操作 |
| `ProofreadService.cs` | 452 | 校对服务（包含并发、缓存、报告功能） |
| `ConfigManager.cs` | 271 | 配置管理，DPAPI 加密存储 |
| `WordDocumentHelper.cs` | 147 | Word COM 操作封装 |
| `BaseLLMService.cs` | 148 | LLM 服务基类 |

### 关键技术
- **并发处理**：支持 3-5 段并行校对（Semaphore 控制）
- **缓存机制**：内存级缓存，基于 SHA256 内容哈希
- **分段处理**：1500 字/段，100 字重叠防止边界错误
- **配置加密**：使用 DPAPI 加密 API Key 和配置

## 使用说明

### 开发环境
1. Visual Studio 2019/2022
2. .NET Framework 4.8
3. Microsoft Office 开发工具

### 依赖包
- `Newtonsoft.Json` - JSON 序列化
- `Microsoft.Office.Interop.Word` - Word 自动化

### 安装步骤
1. 克隆仓库并使用 Visual Studio 打开
2. 还原 NuGet 包
3. 编译项目（Release 或 Debug）
4. 运行 Word，在 COM 加载项中启用 gowordagent

### 使用流程
1. 启动 Word 并打开需要校对的文档
2. 在 Add-In 侧边栏选择 AI 提供商并配置 API Key
3. 点击"保存并连接"测试连接
4. 选择校验模式（精准校验/全文校验）
5. 点击"纠错（审阅）"按钮开始校对
6. 在 Word 审阅面板中接受或拒绝修订建议

## 配置说明

配置文件保存在 `%AppData%\GOWordAgentAddIn\config.dat`，包含：
- AI 提供商和 API 配置
- 自定义提示词（支持两种模式独立配置）
- 自动连接设置

## 注意事项

1. API Key 使用 DPAPI 加密存储，非明文保存
2. 校对前请确保文档已保存
3. 大文档（>3000字）会自动分段处理
4. 缓存仅在当前 Word 会话有效，关闭后清空

## 相关链接

- 原项目地址：https://github.com/jsxyhelu/GOWordAgentAddIn.git
- 飞书文档：https://uh9iow7vir.feishu.cn/wiki/U8Dpwc2NWie8J3kH2FNcahuinab
- 博客文章：https://www.cnblogs.com/jsxyhelu/p/19497787

## 许可证

MIT License

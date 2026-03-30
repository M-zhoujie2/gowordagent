# 更新日志

所有显著变更都会记录在此文件。

格式基于 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.0.0/)。

## [未发布]

### 新增
- 添加 HttpClient 共享工厂，优化连接池管理
- 添加 LLMServiceException 自定义异常，提供更好的错误信息
- 添加 WordProofreadController，分离文档操作逻辑
- 添加 GetDocumentTextEx 方法，支持更精细的文档内容提取
- 添加 DataTemplate 支持，使用 MVVM 模式重构聊天消息渲染

### 优化
- 校对请求超时时间延长至 5 分钟（适配 GLM-4.7 长推理场景）
- 重构 BuildRequestBodyDict 方法，使用 Dictionary 替代匿名对象
- GLM 特有参数（enable_thinking）改为通过覆盖字典方法添加
- 生成报告格式从 Markdown 改为纯文本，适配 WPF TextBlock
- 按钮样式统一到 PrimaryButtonStyle 和 OutlineButtonStyle

### 修复
- 修复 API 返回错误时返回字符串而非抛出异常的问题
- 修复 Word 文档中 ActiveDocument 被错误释放的问题
- 修复 COM 对象未正确释放导致的内存泄漏
- 修复 ProofreadItem 标记 [Obsolete] 但仍被使用的问题

## [1.0.0] - 2024-03-30

### 新增
- 初始版本发布
- 支持 DeepSeek、智谱 GLM、Ollama 三种 AI 提供商
- 精准校验和全文校验两种模式
- Word 原生修订功能集成
- 段落级并发处理（默认 5 并发）
- 内存缓存机制（LRU 淘汰，最大 1000 条）
- 配置加密存储（DPAPI）

### 技术特性
- VSTO Word 插件架构
- .NET Framework 4.8
- WPF 用户界面
- MVVM 数据绑定
- HttpClient 连接池共享
- 完善的 COM 对象生命周期管理

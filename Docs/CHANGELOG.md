# 更新日志

所有显著变更都会记录在此文件。

格式基于 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.0.0/)。

## [1.1.0] - 2024-03-30

### 新增
- 多文档场景支持 - `WordProofreadController` 绑定特定文档，防止切换文档导致引用失效
- SHA256 缓存优化 - 使用静态实例 + 锁，避免高并发时 `SHA256.Create()` 的锁竞争
- DPI 感知配置 - 修复 MaterialDesign 在高分辨率屏幕的模糊问题
- Word 版本检测 - 自动检测 Word 版本，API 降级兼容 Word 2010/2013
- `TryCreateForDocument` 工厂方法 - 支持为特定文档创建服务
- HMAC-SHA256 配置完整性验证 - 防止配置文件篡改

### 优化
- COM 对象生命周期管理 - 统一使用 `try-finally` + `Marshal.ReleaseComObject`
- `SolidColorBrush` 静态复用 - 高频创建的 brush 提取为静态只读字段，减少 GC 压力
- `ScrollIntoView` 添加兼容性保护 - Word 2010 部分版本不支持此方法
- 修订模式降级策略 - `TrackRevisions` 失败时自动降级为普通替换
- `ProofreadService` 优雅关闭 - Dispose 时等待任务完成，避免 `ObjectDisposedException`
- `ApplyRevisionAtRange` 异常处理 - `find` 对象统一在 finally 中释放

### 修复
- 修复 COM 对象泄漏 - `FindTextPosition` 中 `searchRange.Text` 访问可能导致临时对象泄漏
- 修复 `ApplyRevisionAtRange` 中 `find` 对象提前释放的问题
- 修复多文档切换时文档引用失效的问题
- 修复高 DPI 屏幕 UI 模糊的问题
- 修复 Word 2010 不支持 `wdMixedRevisions` 导致的异常

## [1.0.0] - 2024-03-30

### 新增
- 初始版本发布
- 支持 DeepSeek、智谱 GLM、Ollama 三种 AI 提供商
- 精准校验和全文校验两种模式
- Word 原生修订功能集成
- 段落级并发处理（默认 5 并发）
- 内存缓存机制（LRU 淘汰，最大 1000 条）
- 配置加密存储（DPAPI）
- HttpClient 共享工厂，优化连接池管理
- LLMServiceException 自定义异常
- WordProofreadController，分离文档操作逻辑
- DataTemplate 支持，使用 MVVM 模式重构聊天消息渲染

### 技术特性
- VSTO Word 插件架构
- .NET Framework 4.8
- WPF 用户界面
- MVVM 数据绑定
- HttpClient 连接池共享
- 完善的 COM 对象生命周期管理

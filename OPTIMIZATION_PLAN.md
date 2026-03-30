# 代码优化方案

## 一、项目统计

| 文件 | 代码行数 | 说明 |
|------|---------|------|
| GOWordAgentPaneWpf.xaml.cs | 1509 行 | UI 主代码，最臃肿，需重点优化 |
| ProofreadService.cs | 663 行 | 校对服务，逻辑较复杂 |
| BaseLLMService.cs | 287 行 | LLM 基类，带日志功能 |
| ThisAddIn.Designer.cs | 236 行 | VSTO 生成代码，无需改动 |
| ConfigManager.cs | 239 行 | 配置管理 |
| OllamaService.cs | 188 行 | Ollama 服务实现 |
| LLMServiceFactory.cs | 75 行 | 工厂类 |
| ThisAddIn.cs | 72 行 | 插件入口 |
| gowordagentribbon.designer.cs | 83 行 | Ribbon 设计器生成 |
| ILLMService.cs | 36 行 | 接口定义 |
| GLMService.cs | 28 行 | 智谱 AI 服务 |
| gowordagentribbon.cs | 18 行 | Ribbon 代码 |
| DeepSeekService.cs | 16 行 | DeepSeek 服务 |
| GOWordAgentPaneHost.cs | 23 行 | WinForms 宿主 |
| **总计** | **3473 行** | |

## 二、问题分析

### 1. GOWordAgentPaneWpf.xaml.cs (1509 行) - 最严重
**问题：**
- 一个文件包含太多职责：UI 渲染、事件处理、Word 操作、日志记录、状态管理
- `AddMessageBubble` 重复代码多
- 校对结果展示逻辑复杂
- 缺少分层架构

### 2. ProofreadService.cs (663 行)
**问题：**
- 缓存逻辑和分段逻辑混在一起
- 正则表达式硬编码
- 缺少配置化

### 3. 整体架构问题
- 没有 ViewModel 层
- 直接操作 Word 对象，缺少抽象
- 日志分散在各处

## 三、优化方案

### 阶段 1：提取 Word 文档操作层 (预计减少 300 行)

新建 `WordDocumentService.cs`：
```csharp
public class WordDocumentService
{
    // 获取文档文本
    public string GetDocumentText()
    
    // 应用修订
    public void ApplyRevision(string original, string modified, string comment)
    
    // 导航到指定位置
    public void NavigateToRange(int start, int end)
    
    // 查找文本
    public Word.Range FindText(string text)
}
```

### 阶段 2：创建 ViewModel 层 (预计减少 400 行)

新建 `ProofreadResultViewModel.cs`：
```csharp
public class ProofreadResultViewModel
{
    public string Title { get; set; }
    public string ReportContent { get; set; }
    public List<ProofreadIssueItem> Items { get; set; }
    public ICommand CopyCommand { get; }
    public ICommand NavigateCommand { get; }
}
```

### 阶段 3：重构消息气泡 (预计减少 200 行)

新建 `MessageBubbleFactory.cs`：
```csharp
public static class MessageBubbleFactory
{
    public static Border CreateSystemBubble(string message)
    public static Border CreateUserBubble(string message)
    public static Border CreateErrorBubble(string message)
    public static Border CreateProofreadResultBubble(ProofreadResultViewModel vm)
}
```

### 阶段 4：优化 ProofreadService (预计减少 150 行)

- 提取分段逻辑到 `DocumentSegmenter`
- 提取缓存逻辑到 `ProofreadCacheManager`
- 配置化正则表达式

## 四、预期效果

| 文件 | 当前行数 | 优化后 | 减少 |
|------|---------|--------|------|
| GOWordAgentPaneWpf.xaml.cs | 1509 | 800 | 709 |
| ProofreadService.cs | 663 | 450 | 213 |
| 新增文件 | - | 600 | -600 |
| **总计** | **3473** | **2850** | **623** |

## 五、风险点

1. Word 操作需要单元测试
2. 缓存逻辑改动需回归测试
3. UI 重构需保持原有功能

## 六、建议实施顺序

1. 先备份（已完成）
2. 提取 WordDocumentService（低风险）
3. 重构消息气泡（中风险）
4. 创建 ViewModel（高风险，需测试）
5. 优化 ProofreadService（中风险）

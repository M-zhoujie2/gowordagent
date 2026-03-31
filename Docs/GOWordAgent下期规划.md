# GOWordAgent 下期规划

> 整合版本：v1.0 | 日期：2026-03-31
> 包含：WPS COM 兼容改造 + 信创适配改造

---

## 一、规划概述

### 1.1 两个方案的关系

| 方案 | 目标环境 | 改动量 | 工期 | 优先级 |
|------|----------|--------|------|--------|
| **方案一：WPS COM 兼容** | Windows + WPS Pro | ~550 行，~14% 改动 | 7-10 天 | **P0 - 短期** |
| **方案二：信创适配** | Linux + 麒麟/UOS + WPS | ~2000 行，前后端分离 | 6-8 周 | **P1 - 中长期** |

### 1.2 建议执行顺序

```
第一阶段（短期 1-2 周）：WPS COM 兼容改造
    ↓ 验证 WPS COM API 兼容性，让 Windows WPS 用户先用上
    
第二阶段（中长期 6-8 周）：信创适配改造
    ↓ 完整的 Linux 信创环境支持
    
第三阶段（后续）：功能增强
    ↓ 基于新架构扩展功能
```

---

## 二、方案一：WPS COM 兼容改造（Windows 环境）

### 2.1 核心思路

WPS Windows 版暴露了与 Word 兼容的 COM 接口（ProgID: `KWps.Application`），通过以下策略让现有 VSTO 插件最小改动运行在 WPS 上：

1. **不换架构、不拆前后端**，保持 VSTO 单体结构
2. **抽象 Word COM 调用层**，根据宿主自动切换 Word/WPS 实现
3. **替换 VSTO 加载机制为 COM Add-in**（WPS 支持标准 COM Add-in）

### 2.2 WPS COM 能力支持情况

| 能力 | WPS 支持情况 | 本项目是否依赖 |
|------|------------|--------------|
| `Application` 对象 | ✅ `KWps.Application` | ✅ 是 |
| `ActiveDocument` | ✅ | ✅ 是 |
| `Range` 文本操作 | ✅ 基础操作 | ✅ 是 |
| `Range.Find` | ✅ | ✅ 是 |
| `TrackRevisions` | ✅ | ✅ 是 |
| `Comments.Add` | ✅ | ✅ 是 |
| `CustomTaskPane` | ❌ **不支持** | ✅ 是 |
| VSTO Runtime | ❌ **不支持** | ✅ 是 |

**关键障碍**：WPS 不支持 VSTO 的 `CustomTaskPane` 和 `Ribbon` 扩展。

### 2.3 架构改动

```
改动前（Word 专用）：
  VSTO 插件
  ├── ThisAddIn (VSTO 入口)
  ├── GOWordAgentRibbon (VSTO Ribbon)
  ├── GOWordAgentPaneHost (WinForms → WPF 宿主)
  └── WordDocumentService.cs (Word COM)

改动后（Word + WPS 双宿主）：
  COM Add-in（注册表加载）
  ├── ThisAddIn (改为标准 COM Add-in 入口)
  ├── UI 层
  │   ├── Word 版：保持 VSTO Ribbon + CustomTaskPane + WPF
  │   └── WPS 版：独立 WPF Window（浮动窗口替代侧边栏）
  └── Interop/
      ├── IWordInterop.cs（新增接口）
      ├── MsWordInterop.cs（Word 实现）
      └── WpsWordInterop.cs（WPS 实现）
```

### 2.4 新增/改造文件清单

#### 新增文件（7 个）

| 文件 | 说明 |
|------|------|
| `Interop/IWordInterop.cs` | COM 操作抽象接口 |
| `Interop/HostType.cs` | 宿主类型枚举 |
| `Interop/HostDetector.cs` | 宿主检测 |
| `Interop/InteropFactory.cs` | 统一工厂 |
| `Interop/MsWordInterop.cs` | Word COM 实现（原 WordDocumentService 重构） |
| `Interop/WpsWordInterop.cs` | WPS COM 适配实现 |
| `WpfFloatingWindow.cs` | WPS 浮动窗口（替代 CustomTaskPane） |

#### 改造文件（3 个）

| 文件 | 改动内容 |
|------|----------|
| `ThisAddIn.cs` | 增加宿主检测和 WPS 分支初始化 |
| `WordProofreadController.cs` | WordDocumentService → IWordInterop（~20 行改动） |
| `gowordagentribbon.cs` | 增加 WPS 窗口切换分支（~5 行改动） |

#### 无需改动文件（17 个）

`GOWordAgentPaneWpf.xaml/cs`, `ProofreadService.cs`, `ProofreadCacheManager.cs`, `ConfigManager.cs`, `ILLMService.cs`, `DeepSeekService.cs`, `GLMService.cs`, `OllamaService.cs`, `LLMServiceFactory.cs`, `HttpClientFactory.cs`, `LLMRequestLogger.cs`, `ProofreadIssueParser.cs`, `DocumentSegmenter.cs`, `ProofreadResultRenderer.cs`, `ViewModels/*.cs`, `Models/ProofreadModels.cs`

### 2.5 工期估算（7-10 天）

| 阶段 | 工作项 | 工期 | 前置依赖 |
|------|--------|------|---------|
| **S0** | WPS COM 兼容性验证（运行测试脚本） | **1-2 天** | 需要 Windows + WPS Pro |
| **S1** | 抽象接口 IWordInterop + MsWordInterop | **1 天** | S0 |
| **S2** | WpsWordInterop 实现 + InteropFactory | **2 天** | S0 + S1 |
| **S3** | WordProofreadController 改造 | **半天** | S1 + S2 |
| **S4** | WpfFloatingWindow + ThisAddIn WPS 分支 | **1 天** | S2 |
| **S5** | Ribbon 适配 + COM Add-in 注册表 | **1 天** | S4 |
| **S6** | 端到端测试（Word + WPS 双环境） | **2 天** | S5 |

### 2.6 风险与降级

| 风险 | 概率 | 影响 | 降级方案 |
|------|------|------|---------|
| `TrackRevisions` 在 WPS 不生效 | 中 | 高 | 降级为直接替换+批注 |
| `Comments.Add` 参数顺序不同 | 中 | 中 | 尝试两种参数顺序 |
| WPS 版本间 COM API 不一致 | 中 | 高 | 针对目标版本逐个测试 |

---

## 三、方案二：信创适配改造（Linux 环境）

### 3.1 现状评估

通过源码分析，各文件与 Windows/Word 的耦合程度：

| 类型 | 文件数 | 行数 | 可复用度 |
|------|--------|------|---------|
| **可直接复用** | ~10 | ~1160 行 | 100% |
| **需少量改动复用** | ~2 | ~400 行 | 60-90% |
| **需重写** | ~8 | ~2350 行 | 0% |

**实际复用率**：约 30%（按行计），约 50%（按逻辑价值计）

### 3.2 改造后架构

```
┌─────────────────────────────────────────────┐
│  WPS JS 插件（前端）                         │
│  ┌──────────┐  ┌──────────┐  ┌───────────┐ │
│  │ 侧边栏UI │  │ 文档操作 │  │ 配置面板   │ │
│  │ (HTML)   │  │(WPS JS   │  │ (HTML)    │ │
│  └────┬─────┘  │  API)    │  └─────┬─────┘ │
│       └──────────┬──────────────┘        │
│            HTTP / WebSocket                │
├─────────────────────────────────────────────┤
│  本地后端服务（ASP.NET Core）                │
│  ┌──────────┐  ┌──────────┐  ┌───────────┐ │
│  │ Proofread│  │ Config   │  │ LLM       │ │
│  │ Service  │  │ Manager  │  │ Services  │ │
│  │ (复用+解耦)│  │(重写加密)│  │ (直接复用) │ │
│  └──────────┘  └──────────┘  └───────────┘ │
├─────────────────────────────────────────────┤
│  LLM API（DeepSeek/GLM/Ollama）             │
└─────────────────────────────────────────────┘
```

### 3.3 后端改造（ASP.NET Core）

#### 项目结构

```
gowordagent-server/
├── Controllers/
│   └── ProofreadController.cs    # HTTP API 端点
├── Services/                      # 从原项目复制 + 改造
│   ├── ProofreadService.cs       # 解耦 Dispatcher
│   ├── DeepSeekService.cs        # 直接复制
│   ├── GLMService.cs             # 直接复制
│   ├── OllamaService.cs          # 直接复制
│   ├── LLMServiceFactory.cs      # 直接复制
│   ├── ProofreadCacheManager.cs  # 改路径为跨平台
│   └── DocumentSegmenter.cs      # 直接复制
├── Infrastructure/
│   ├── WsProgressReporter.cs     # WebSocket 进度推送
│   └── CryptoService.cs          # AES 替代 DPAPI
└── Models/
    └── ProofreadModels.cs        # 直接复制
```

#### 关键改造点

**1. ProofreadService 解耦**

```csharp
// 新增接口
public interface IProgressReporter
{
    Task ReportProgressAsync(ProofreadProgressArgs args);
}

// 改造后：删除 WPF Dispatcher 依赖
public ProofreadService(
    ILLMService llmService,
    string systemPrompt,
    int concurrency,
    IProgressReporter progressReporter)  // 替换 Dispatcher
{
    _progressReporter = progressReporter;
}
```

**2. ConfigManager 加密层替换（DPAPI → AES）**

```csharp
public static class CryptoService
{
    // AES-256 替代 Windows DPAPI
    public static byte[] Encrypt(string plainText)
    public static string Decrypt(byte[] data)
    
    // 密钥存储：~/.config/gowordagent/crypto.key
    // Linux 权限：chmod 600
}
```

### 3.4 前端改造（WPS JS 插件）

#### 目录结构

```
gowordagent-wps/
├── plugin.json                    # WPS 插件清单
├── sidebar/
│   ├── index.html                 # 侧边栏主页面
│   ├── css/sidebar.css            # 样式
│   └── js/
│       ├── sidebar.js             # UI 交互逻辑
│       ├── api-client.js          # 调用后端 HTTP API
│       ├── document-service.js    # WPS JS 文档操作
│       └── proofread-controller.js# 修订/批注控制
└── assets/icons/
```

#### WPS JS API 兼容性验证清单

| API | 优先级 | 风险等级 |
|-----|--------|---------|
| `Application.ActiveDocument.Content.Text` | P0 | 低 |
| `Application.ActiveWindow.Selection.Text` | P0 | 低 |
| `range.Find.Execute(...)` | P0 | 中 |
| `document.TrackRevisions = true` | P1 | **高** |
| `document.Comments.Add(range, text)` | P1 | 中 |
| `Application.ActiveWindow.ScrollIntoView` | P2 | 低（有降级） |

### 3.5 进程生命周期管理

**systemd 用户服务（麒麟/UOS）**

```ini
# ~/.config/systemd/user/gowordagent.service
[Unit]
Description=GOWordAgent Proofread Service

[Service]
Type=simple
ExecStart=/opt/gowordagent/gowordagent-server
Restart=on-failure
Environment=ASPNETCORE_URLS=http://127.0.0.1:19527

[Install]
WantedBy=default.target
```

### 3.6 工期估算（6-8 周）

| 阶段 | 工作项 | 工期 | 前置依赖 |
|------|--------|------|---------|
| **S0: POC** | WPS JS API 兼容性验证 + 目标环境实测 | **3-5 天** | 需要麒麟/UOS + WPS |
| **S1: 后端骨架** | ASP.NET Core 项目 + ProofreadService 解耦 + HTTP API | **1 周** | 无 |
| **S2: 后端完善** | WebSocket 进度 + 配置 API + 跨平台编译 | **1 周** | S1 |
| **S3: 前端 UI** | 侧边栏 HTML/CSS + 配置面板 + 消息列表 | **1.5 周** | S2 |
| **S4: 前端文档操作** | document-service.js + 逐 API 验证 | **1.5 周** | S0 + S3 |
| **S5: 集成调试** | 端到端联调 + 修复 WPS API 差异 | **1 周** | S2 + S4 |
| **S6: 测试** | 麒麟 x86 + UOS ARM64 全流程测试 | **1 周** | S5 |
| **合计** | | **6-8 周** | |

---

## 四、两方案对比总结

| 维度 | 方案一（WPS COM） | 方案二（信创适配） |
|------|-------------------|-------------------|
| **改动量** | ~550 行 | ~2000 行 |
| **工期** | 7-10 天 | 6-8 周 |
| **代码复用率** | ~85% | ~30% |
| **运行环境** | Windows only | Windows + Linux |
| **信创支持** | ❌ 不支持 | ✅ 支持 |
| **架构变化** | 单体 VSTO | 前后端分离 |
| **部署方式** | COM Add-in | systemd + WPS JS 插件 |

---

## 五、实施建议

### 5.1 推荐路线

1. **先执行方案一**（WPS COM 兼容）：
   - 投入小（7-10 天），产出快
   - 让 Windows WPS 用户立即受益
   - 验证 WPS API 兼容性，为方案二积累经验

2. **后执行方案二**（信创适配）：
   - 长期布局，支持国家信创战略
   - 基于方案一已验证的 WPS API 行为，降低风险

### 5.2 POC 验证要点

在正式开始开发前，务必完成以下 POC 验证：

**方案一 POC（1-2 天）**：
- [ ] `TrackRevisions` 在 WPS 下是否生效
- [ ] `Comments.Add` 参数顺序
- [ ] `Find.Execute` Wrap 参数数值

**方案二 POC（3-5 天）**：
- [ ] 能获取文档全文和选中文本
- [ ] 能调用后端 LLM API 并返回校对结果
- [ ] 能用修订模式写入文档
- [ ] 能添加批注
- [ ] WebSocket 进度推送正常

### 5.3 文件改动总览

```
Docs/
├── GOWordAgent-WPS-COM兼容改造方案.md      # 详细技术方案一
├── GOWordAgent信创适配改造方案.md           # 详细技术方案二
└── GOWordAgent下期规划.md                   # 本文件（整合规划）
```

---

*本规划基于 gowordagent 源码实际分析，建议先完成 POC 验证后再推进后续开发工作。*

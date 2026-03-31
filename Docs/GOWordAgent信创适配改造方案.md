# GOWordAgent 信创适配改造方案

> 版本：v1.0 | 日期：2026-03-31 | 基于源码分析

---

## 一、现状精确评估

### 1.1 源码耦合分析

通过实际阅读源码，各文件与 Windows/Word 的耦合程度如下：

| 文件 | 行数 | Windows 耦合 | 可复用度 | 说明 |
|------|------|-------------|---------|------|
| `Models/ProofreadModels.cs` | ~80 | ❌ 无 | **100%** | 纯数据模型，直接复用 |
| `ILLMService.cs` | 接口 | ❌ 无 | **100%** | 纯接口定义 |
| `DeepSeekService.cs` | ~150 | ❌ 无 | **100%** | HttpClient 调用，无平台依赖 |
| `GLMService.cs` | ~150 | ❌ 无 | **100%** | 同上 |
| `OllamaService.cs` | ~150 | ❌ 无 | **100%** | 同上 |
| `LLMServiceFactory.cs` | ~80 | ❌ 无 | **100%** | 工厂模式，无平台依赖 |
| `HttpClientFactory.cs` | ~50 | ❌ 无 | **100%** | HTTP 客户端工厂 |
| `LLMRequestLogger.cs` | ~80 | ❌ 无 | **100%** | Debug 日志 |
| `ProofreadIssueParser.cs` | ~200 | ❌ 无 | **95%** | 正则解析 LLM 返回，纯字符串操作 |
| `DocumentSegmenter.cs` | ~100 | ❌ 无 | **95%** | 文本分段逻辑 |
| `ProofreadCacheManager.cs` | ~100 | ❌ 无 | **90%** | SHA256+内存缓存，需改路径 |
| `ConfigManager.cs` | ~280 | ⚠️ DPAPI | **30%** | DPAPI 加密 + Windows 路径，需重写加密层 |
| `ProofreadService.cs` | ~300 | ⚠️ WPF Dispatcher | **60%** | 核心逻辑可复用，但通过 `Dispatcher.InvokeAsync` 报告进度，需改为回调接口 |
| `WordDocumentService.cs` | ~500 | 🔴 Word COM | **0%** | 全部是 `Microsoft.Office.Interop.Word` 操作，需用 WPS JS API 完全重写 |
| `WordProofreadController.cs` | ~200 | 🔴 Word COM | **0%** | 调用 WordDocumentService，需重写 |
| `GOWordAgentPaneWpf.xaml.cs` | ~1100 | 🔴 WPF | **0%** | WPF 控件、XAML 绑定、Word 事件，需完全重写 |
| `ViewModels/*.cs` | ~200 | ⚠️ WPF | **0%** | WPF 数据绑定相关 |
| `ProofreadResultRenderer.cs` | ~150 | ⚠️ WPF | **0%** | WPF UI 渲染 |
| `gowordagentribbon.cs` | ~100 | 🔴 Word Ribbon | **0%** | Word 自定义 Ribbon |
| `ThisAddIn.cs` | ~50 | 🔴 VSTO | **0%** | VSTO 入口 |

### 1.2 复用度精确结论

```
可直接复用（零改动）：  ~1160 行（LLM 服务层 + 数据模型 + 工具类）
需少量改动复用：         ~400 行（ProofreadService 解耦、CacheManager 路径）
需重写：               ~2350 行（UI、文档操作、配置加密）
合计：                 ~3910 行
实际复用率：           ~30%（按行计），~50%（按逻辑价值计）
```

### 1.3 ProofreadService 耦合点详解

ProofreadService 是核心服务，有 3 个 Windows 耦合点需要解耦：

```csharp
// 耦合点 1：构造函数中获取 WPF Dispatcher
_dispatcher = Application.Current?.Dispatcher ?? Dispatcher.CurrentDispatcher;

// 耦合点 2：通过 Dispatcher 回调报告进度
await _dispatcher.InvokeAsync(() => ReportProgress(...));

// 耦合点 3：Debug.WriteLine（非 Windows 耦合，但要换日志框架）
Debug.WriteLine($"[ProofreadService] ...");
```

**解耦方案**：把 `Dispatcher` 替换为 `SynchronizationContext` 或纯回调接口：

```csharp
// 改造后
public interface IProgressReporter
{
    void ReportProgress(ProofreadProgressArgs args);
}

public ProofreadService(ILLMService llmService, string systemPrompt,
    int concurrency, IProgressReporter progressReporter)
{
    _progressReporter = progressReporter;
    // 不再依赖 WPF Dispatcher
}
```

---

## 二、改造方案

### 2.1 整体架构

```
┌─────────────────────────────────────────────┐
│  WPS JS 插件（前端）                         │
│  ┌──────────┐  ┌──────────┐  ┌───────────┐ │
│  │ 侧边栏UI │  │ 文档操作 │  │ 配置面板   │ │
│  │ (HTML)   │  │(WPS JS   │  │ (HTML)    │ │
│  │          │  │  API)    │  │           │ │
│  └────┬─────┘  └──────────┘  └─────┬─────┘ │
│       └──────────┬─────────────────┘        │
│            HTTP / WebSocket                   │
├─────────────────────────────────────────────┤
│  本地后端服务（ASP.NET Core）                 │
│  ┌──────────┐  ┌──────────┐  ┌───────────┐ │
│  │ Proofread│  │ Config   │  │ LLM       │ │
│  │ Service  │  │ Manager  │  │ Services  │ │
│  │ (复用+解耦)│  │(重写加密)│  │ (直接复用) │ │
│  └──────────┘  └──────────┘  └───────────┘ │
├─────────────────────────────────────────────┤
│  LLM API（DeepSeek/GLM/Ollama）              │
└─────────────────────────────────────────────┘
```

### 2.2 后端改造（C# → ASP.NET Core）

#### Step 1：新建 ASP.NET Core Web API 项目

```
gowordagent-server/
├── gowordagent-server.csproj     # .NET 8, net8.0 / net8.0-linux-arm64
├── Program.cs
├── Controllers/
│   └── ProofreadController.cs    # HTTP API 端点
├── Services/                      # 从原项目复制 + 改造
│   ├── ProofreadService.cs       # 解耦 Dispatcher，改为 IProgressReporter
│   ├── DeepSeekService.cs        # 直接复制
│   ├── GLMService.cs             # 直接复制
│   ├── OllamaService.cs          # 直接复制
│   ├── LLMServiceFactory.cs      # 直接复制
│   ├── HttpClientFactory.cs      # 直接复制
│   ├── ProofreadIssueParser.cs   # 直接复制
│   ├── DocumentSegmenter.cs      # 直接复制
│   └── ProofreadCacheManager.cs  # 改路径为跨平台
├── Models/                        # 直接复制
│   └── ProofreadModels.cs
├── Config/
│   └── ConfigManager.cs          # 重写加密层
└── Infrastructure/
    ├── WsProgressReporter.cs     # WebSocket 进度推送
    └── CryptoService.cs          # AES 替代 DPAPI
```

#### Step 2：ProofreadService 解耦

```csharp
// 新增接口（新增文件）
public interface IProgressReporter
{
    Task ReportProgressAsync(ProofreadProgressArgs args);
}

// 改造 ProofreadService.cs
public class ProofreadService : IDisposable
{
    private readonly IProgressReporter _progressReporter;

    public ProofreadService(
        ILLMService llmService,
        string systemPrompt,
        int concurrency,
        IProgressReporter progressReporter)  // 替换 Dispatcher
    {
        _progressReporter = progressReporter;
        // 删除: _dispatcher = Application.Current.Dispatcher
    }

    private async Task<ParagraphResult> ProcessParagraphAsync(...)
    {
        // 改前: await _dispatcher.InvokeAsync(() => ReportProgress(...));
        // 改后:
        await _progressReporter.ReportProgressAsync(new ProofreadProgressArgs
        {
            TotalParagraphs = total,
            CompletedParagraphs = completed,
            CurrentIndex = index,
            CurrentStatus = $"正在校对第 {index + 1}/{total} 段...",
            Result = result,
            IsCompleted = false
        });
    }
}
```

#### Step 3：ConfigManager 加密层替换

```csharp
public static class CryptoService
{
    private static readonly string KeyDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
        ".config", "gowordagent");

    private static readonly string KeyFile = Path.Combine(KeyDir, "crypto.key");

    /// <summary>
    /// 获取或生成 AES 密钥（替代 DPAPI）
    /// </summary>
    public static byte[] GetOrCreateKey()
    {
        if (File.Exists(KeyFile))
        {
            return Convert.FromBase64String(File.ReadAllText(KeyFile));
        }
        byte[] key = RandomNumberGenerator.GetBytes(32); // AES-256
        Directory.CreateDirectory(KeyDir);
        File.WriteAllText(KeyFile, Convert.ToBase64String(key));
        // 设置文件权限：仅当前用户可读
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            SetUnixFilePermissions(KeyFile);
        }
        return key;
    }

    public static byte[] Encrypt(string plainText)
    {
        using var aes = Aes.Create();
        aes.Key = GetOrCreateKey();
        aes.GenerateIV();
        using var encryptor = aes.CreateEncryptor();
        byte[] encrypted = encryptor.TransformFinalBlock(
            Encoding.UTF8.GetBytes(plainText), 0, plainText.Length);
        // IV + ciphertext
        return aes.IV.Concat(encrypted).ToArray();
    }

    public static string Decrypt(byte[] data)
    {
        using var aes = Aes.Create();
        aes.Key = GetOrCreateKey();
        aes.IV = data[..16]; // 前 16 字节是 IV
        using var decryptor = aes.CreateDecryptor();
        byte[] decrypted = decryptor.TransformFinalBlock(
            data, 16, data.Length - 16);
        return Encoding.UTF8.GetString(decrypted);
    }

    private static void SetUnixFilePermissions(string path)
    {
        try
        {
            // chmod 600
            var mono = System.Diagnostics.Process.Start("chmod", "600 " + path);
            mono?.WaitForExit(5000);
        }
        catch { }
    }
}
```

#### Step 4：HTTP API 端点

```csharp
[ApiController]
[Route("api/[controller]")]
public class ProofreadController : ControllerBase
{
    private readonly ILLMServiceFactory _llmFactory;

    [HttpPost("start")]
    public async Task<IActionResult> StartProofread([FromBody] ProofreadRequest req)
    {
        var llm = _llmFactory.Create(req.Provider, req.ApiKey, req.ApiUrl, req.Model);
        var reporter = new WebSocketProgressReporter(HttpContext);
        var service = new ProofreadService(llm, prompt, concurrency, reporter);

        _ = Task.Run(() => service.ProofreadDocumentAsync(req.Text, HttpContext.RequestAborted));
        return Accepted(new { taskId = reporter.TaskId });
    }

    [HttpPost("incremental")]
    public async Task<IActionResult> IncrementalProofread([FromBody] ProofreadRequest req)
    {
        // 增量校对：对比已缓存结果
    }

    [HttpGet("cache/stats")]
    public IActionResult GetCacheStats() { ... }

    [HttpDelete("cache")]
    public IActionResult ClearCache() { ... }

    [HttpPost("config")]
    public IActionResult SaveConfig([FromBody] AIConfig config) { ... }

    [HttpGet("config")]
    public IActionResult GetConfig() { ... }

    [HttpPost("test-connection")]
    public async Task<IActionResult> TestConnection([FromBody] ProviderConfig config)
    {
        // 测试 LLM 连通性
    }
}
```

#### Step 5：自包含发布到 Linux

```bash
# x86_64
dotnet publish -c Release -r linux-x64 --self-contained true -o publish/linux-x64

# ARM64（飞腾/鲲鹏）
dotnet publish -c Release -r linux-arm64 --self-contained true -o publish/linux-arm64
```

### 2.3 前端改造（WPS JS 插件）

#### WPS 插件目录结构

```
gowordagent-wps/
├── plugin.json                    # WPS 插件清单
├── sidebar/
│   ├── index.html                 # 侧边栏主页面
│   ├── css/
│   │   └── sidebar.css            # 样式（从 WPF 色值迁移）
│   └── js/
│       ├── sidebar.js             # UI 交互逻辑
│       ├── api-client.js          # 调用后端 HTTP API
│       ├── document-service.js    # WPS JS 文档操作
│       └── proofread-controller.js# 修订/批注控制
└── assets/
    └── icons/
```

#### plugin.json（WPS 插件清单）

```json
{
  "name": "gowordagent-proofread",
  "displayName": "智能校对",
  "version": "1.0.0",
  "minWpsVersion": "11.8.2",
  "main": "sidebar/index.html",
  "description": "AI 智能校对插件，支持精准校验和全文校验",
  "permissions": [
    "document.read",
    "document.write",
    "document.revisions",
    "document.comments"
  ]
}
```

#### document-service.js（核心：Word COM → WPS JS）

```javascript
/**
 * WPS JS 文档操作服务
 * 对应原 WordDocumentService.cs 的功能
 *
 * ⚠️ 每个 API 都需要在目标环境实测验证！
 */

class WpsDocumentService {

  /**
   * 获取文档全文或选中文本
   * 对应: WordDocumentService.GetDocumentText()
   */
  static getText() {
    const doc = Application.ActiveDocument;
    if (!doc) throw new Error('未打开文档');

    const sel = Application.ActiveWindow.Selection;
    if (sel && sel.Text && sel.Text.trim()) {
      return sel.Text;
    }
    return doc.Content.Text;
  }

  /**
   * 查找文本位置（三级匹配策略）
   * 对应: WordDocumentService.FindTextPosition()
   *
   * ⚠️ WPS JS 的 Find API 行为可能与 Word COM 不一致，需实测：
   * - MatchCase 参数名是否一致
   * - MatchWholeWord 是否支持
   * - 返回值的 Start/End 是否与 Word 一致
   */
  static findTextPosition(text) {
    if (!text) return { found: false };

    const range = Application.ActiveDocument.Content;
    const find = range.Find;

    // 第 1 步：精确匹配
    find.ClearFormatting();
    let found = find.Execute(text, true, true); // MatchCase, MatchWholeWord
    if (found) return { found: true, start: range.Start, end: range.End };

    // 第 2 步：不区分大小写 + 整词
    found = find.Execute(text, false, true);
    if (found) return { found: true, start: range.Start, end: range.End };

    // 第 3 步：长文本宽松匹配
    if (text.length > 5) {
      found = find.Execute(text, false, false);
      if (found) return { found: true, start: range.Start, end: range.End };
    }

    return { found: false };
  }

  /**
   * 应用单个修订
   * 对应: WordDocumentService.ApplyRevision()
   *
   * ⚠️ WPS 修订 API 需实测：
   * - TrackRevisions 开关是否生效
   * - 修订后文本替换是否正确触发修订标记
   * - Comments.Add 参数顺序（WPS 可能是 Add(range, text) 或 Add(text, range)）
   */
  static applyRevision(original, modified, commentText) {
    const doc = Application.ActiveDocument;
    const result = this.findTextPosition(original);
    if (!result.found) return false;

    const range = doc.Range(result.start, result.end);

    // 开启修订模式
    const oldTrack = doc.TrackRevisions;
    try {
      doc.TrackRevisions = true;
      range.Text = modified;

      // ⚠️ WPS JS 的 Comments.Add API 需要实测参数顺序
      try {
        doc.Comments.Add(range, commentText);
      } catch (e) {
        console.error('添加批注失败:', e.message);
        // 批注失败不影响修订
      }
      return true;
    } catch (e) {
      console.error('修订模式失败，降级处理:', e.message);
      // 降级：直接替换 + 标记原文
      range.Text = `[原文：${original}] ${modified}`;
      doc.Comments.Add(range, `${commentText}\n⚠️ 修订模式不可用，已直接替换`);
      return true;
    } finally {
      try { doc.TrackRevisions = oldTrack; } catch {}
    }
  }

  /**
   * 批量应用修订（倒序处理）
   * 对应: WordDocumentService.ApplyRevisions()
   */
  static applyRevisions(items) {
    // 先查所有位置
    const positions = items.map(item => {
      const pos = this.findTextPosition(item.original);
      return { item, ...pos };
    }).filter(p => p.found);

    // 倒序排列
    positions.sort((a, b) => b.start - a.start);

    // 逐个应用
    const applied = [];
    for (const { item, start, end } of positions) {
      const comment = [
        `【第${item.index}处】类型：${item.type}`,
        `原文：${item.original}`,
        `修改：${item.modified}`,
        `理由：${item.reason}`
      ].join('\n');

      if (this.applyRevision(item.original, item.modified, comment)) {
        applied.push(item);
      }
    }
    return applied;
  }

  /**
   * 导航到指定位置
   * 对应: WordDocumentService.NavigateToRange()
   */
  static navigateToRange(start, end) {
    const range = Application.ActiveDocument.Range(start, end);
    range.Select();
    // ⚠️ WPS 是否支持 ScrollIntoView？
    try {
      Application.ActiveWindow.ScrollIntoView(range);
    } catch (e) {
      console.warn('ScrollIntoView 不支持:', e.message);
    }
  }

  /**
   * 通过原文搜索定位
   * 对应: WordDocumentService.NavigateBySearch()
   */
  static navigateByText(originalText) {
    const pos = this.findTextPosition(originalText);
    if (pos.found) {
      this.navigateToRange(pos.start, pos.end);
      return true;
    }
    return false;
  }
}
```

#### sidebar.js（前端主逻辑）

```javascript
/**
 * 侧边栏主控制器
 * 对应原 GOWordAgentPaneWpf.xaml.cs
 */

const API_BASE = 'http://localhost:19527/api';

class SidebarController {
  constructor() {
    this.provider = 'DeepSeek';
    this.apiKey = '';
    this.apiUrl = '';
    this.model = '';
    this.proofreadMode = '精准校验';
    this.isProofreading = false;
    this.init();
  }

  async init() {
    // 加载配置
    await this.loadConfig();

    // 绑定事件
    document.getElementById('btn-save').onclick = () => this.saveConfig();
    document.getElementById('btn-connect').onclick = () => this.testConnection();
    document.getElementById('btn-proofread').onclick = () => this.startProofread();

    // WebSocket 连接（接收进度）
    this.connectWebSocket();
  }

  async loadConfig() {
    const resp = await fetch(`${API_BASE}/config`);
    const config = await resp.json();
    this.provider = config.provider;
    this.proofreadMode = config.proofreadMode || '精准校验';
    // 更新 UI...
  }

  async startProofread() {
    if (this.isProofreading) return;
    this.isProofreading = true;

    try {
      // 1. 获取文档文本
      const text = WpsDocumentService.getText();
      if (!text.trim()) {
        this.addMessage('系统', '文档内容为空');
        return;
      }

      // 2. 调用后端开始校对
      const resp = await fetch(`${API_BASE}/proofread/start`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          text,
          provider: this.provider,
          apiKey: this.apiKey,
          apiUrl: this.apiUrl,
          model: this.model,
          mode: this.proofreadMode
        })
      });

      if (resp.status === 202) {
        this.addMessage('系统', '校对已开始，请稍候...');
        // 进度通过 WebSocket 推送
      }
    } catch (e) {
      this.addMessage('系统', `校对失败: ${e.message}`);
    } finally {
      this.isProofreading = false;
    }
  }

  /**
   * 处理校对完成（WebSocket 回调）
   */
  onProofreadCompleted(results) {
    // 将结果渲染为列表（对应原 ProofreadResultRenderer）
    const issues = [];
    results.forEach(r => {
      if (r.items) issues.push(...r.items);
    });

    if (issues.length === 0) {
      this.addMessage('AI', '未发现错误 ✅');
      return;
    }

    this.addMessage('AI', `共发现 ${issues.length} 处问题`);
    // 渲染问题列表，每个问题带"定位"和"应用"按钮
    this.renderIssueList(issues);
  }

  /**
   * 应用所有修订到文档
   */
  async applyAllRevisions(issues) {
    const applied = WpsDocumentService.applyRevisions(issues);
    this.addMessage('系统', `已应用 ${applied.length}/${issues.length} 处修订`);
  }

  /**
   * 跳转到文档对应位置
   */
  navigateToIssue(issue) {
    WpsDocumentService.navigateByText(issue.original);
  }

  connectWebSocket() {
    const ws = new WebSocket('ws://localhost:19527/ws/progress');
    ws.onmessage = (event) => {
      const msg = JSON.parse(event.data);
      if (msg.type === 'progress') {
        this.updateProgress(msg.data);
      } else if (msg.type === 'completed') {
        this.onProofreadCompleted(msg.data.results);
      }
    };
    ws.onclose = () => {
      // 3 秒后重连
      setTimeout(() => this.connectWebSocket(), 3000);
    };
  }

  addMessage(role, text) {
    // 渲染消息气泡到侧边栏
  }
}
```

---

## 三、进程生命周期管理（最大隐性成本）

### 3.1 问题

后端 .NET HTTP 服务需要随 WPS 启动/退出，这是最容易翻车的环节：

| 问题 | 说明 |
|------|------|
| 谁启动服务？ | WPS 插件 JS 无法直接启动 .NET 进程 |
| 服务崩溃怎么办？ | 需要自动重启 |
| 多 WPS 实例冲突？ | 端口占用 |
| 用户怎么装？ | 不能让用户手动 dotnet run |

### 3.2 方案：systemd 用户服务（麒麟/UOS）

```ini
# ~/.config/systemd/user/gowordagent.service
[Unit]
Description=GOWordAgent Proofread Service
After=network.target

[Service]
Type=simple
ExecStart=/opt/gowordagent/gowordagent-server
Restart=on-failure
RestartSec=3
Environment=ASPNETCORE_URLS=http://127.0.0.1:19527

[Install]
WantedBy=default.target
```

```bash
# 安装
systemctl --user enable gowordagent
systemctl --user start gowordagent

# 查看状态
systemctl --user status gowordagent
journalctl --user -u gowordagent -f
```

### 3.3 安装脚本

```bash
#!/bin/bash
# install.sh - 一键安装脚本

set -e

# 检测架构
ARCH=$(uname -m)
case $ARCH in
    x86_64)  BIN_DIR="linux-x64" ;;
    aarch64) BIN_DIR="linux-arm64" ;;
    *) echo "不支持的架构: $ARCH"; exit 1 ;;
esac

# 安装后端
INSTALL_DIR="/opt/gowordagent"
sudo mkdir -p $INSTALL_DIR
sudo cp -r bin/$BIN_DIR/* $INSTALL_DIR/
sudo chmod +x $INSTALL_DIR/gowordagent-server

# 安装 systemd 服务
mkdir -p ~/.config/systemd/user
cp gowordagent.service ~/.config/systemd/user/
systemctl --user daemon-reload
systemctl --user enable gowordagent
systemctl --user start gowordagent

# 安装 WPS 插件
WPS_ADDON_DIR="/opt/kingsoft/wps-office/office6/addons/gowordagent"
sudo mkdir -p $WPS_ADDON_DIR
sudo cp -r wps-plugin/* $WPS_ADDON_DIR/

# 验证
echo "等待服务启动..."
sleep 2
if curl -s http://127.0.0.1:19527/api/config > /dev/null; then
    echo "✅ 安装成功！请重启 WPS 以加载插件。"
else
    echo "❌ 服务启动失败，请检查日志："
    journalctl --user -u gowordagent --no-pager -n 20
fi
```

---

## 四、WPS JS API 兼容性验证清单

**这是风险最大的部分，必须在目标环境逐项实测。**

### 4.1 必须验证的 API（按优先级排序）

```javascript
// ===== P0：没有这些就无法工作 =====

// 1. 获取文档内容
Application.ActiveDocument.Content.Text
// 预期：返回文档全文
// 风险：表格/图片/特殊字符处理可能不同

// 2. 获取选中文本
Application.ActiveWindow.Selection.Text
// 预期：返回选中文本
// 风险：Selection 对象属性可能不同

// 3. 文本查找
range.Find.Execute(FindText, MatchCase, MatchWholeWord)
// 预期：与 Word COM 行为一致
// 风险：参数名/顺序可能不同，返回值结构可能不同

// 4. 文本替换
range.Text = "new text"
// 预期：替换 range 内容
// 风险：低，基础 API 通常兼容

// ===== P1：修订功能核心 =====

// 5. 修订模式开关
document.TrackRevisions = true
// 预期：开启修订模式
// 风险：⚠️ 中，WPS 可能行为不同

// 6. 获取修订内容
range.Text（在 TrackRevisions=true 时替换）
// 预期：自动产生修订标记
// 风险：⚠️ 高，这是最关键的差异点

// 7. 添加批注
document.Comments.Add(range, text)
// 预期：在指定位置添加批注
// 风险：⚠️ 中，参数顺序可能不同

// ===== P2：辅助功能 =====

// 8. 选区导航
range.Select()
Application.ActiveWindow.ScrollIntoView(range)
// 预期：选中并滚动到目标位置
// 风险：ScrollIntoView 可能不支持

// 9. 段落遍历
document.Paragraphs.Count
document.Paragraphs(i).Range.Text
// 预期：遍历段落
// 风险：低

// 10. 范围定位
document.Range(start, end)
// 预期：创建指定范围
// 风险：低
```

### 4.2 验证脚本

```javascript
// compat-test.js - 在 WPS 插件面板中运行
// 每项测试结果输出到面板和控制台

const tests = [
  {
    name: '获取文档内容',
    fn: () => Application.ActiveDocument.Content.Text
  },
  {
    name: '获取选中文本',
    fn: () => Application.ActiveWindow?.Selection?.Text || '无选中'
  },
  {
    name: '段落数量',
    fn: () => Application.ActiveDocument.Paragraphs.Count
  },
  {
    name: 'TrackRevisions 开关',
    fn: () => {
      const doc = Application.ActiveDocument;
      const old = doc.TrackRevisions;
      doc.TrackRevisions = true;
      const after = doc.TrackRevisions;
      doc.TrackRevisions = old;
      return { old, set: true, after };
    }
  },
  {
    name: 'Comments.Add',
    fn: () => {
      const doc = Application.ActiveDocument;
      const range = doc.Range(0, 5);
      try {
        const comment = doc.Comments.Add(range, '测试批注');
        comment.Delete(); // 立即删除
        return '✅ 成功';
      } catch (e) {
        return `❌ ${e.message}`;
      }
    }
  },
  {
    name: 'Find.Execute',
    fn: () => {
      const range = Application.ActiveDocument.Content;
      const found = range.Find.Execute('的', false, false);
      return found ? `✅ 找到, pos=${range.Start}-${range.End}` : '未找到';
    }
  },
  {
    name: 'Range.Select + ScrollIntoView',
    fn: () => {
      const range = Application.ActiveDocument.Range(0, 10);
      range.Select();
      try {
        Application.ActiveWindow.ScrollIntoView(range);
        return '✅ 成功';
      } catch (e) {
        return `⚠️ ScrollIntoView 失败: ${e.message}`;
      }
    }
  }
];

async function runTests() {
  const results = [];
  for (const test of tests) {
    try {
      const result = test.fn();
      results.push({ name: test.name, status: '✅', result: JSON.stringify(result) });
    } catch (e) {
      results.push({ name: test.name, status: '❌', result: e.message });
    }
  }
  console.table(results);
  return results;
}
```

---

## 五、文件改动清单（精确到文件）

### 5.1 直接复用（复制，不改或微改）

| 原文件 | 操作 | 改动 |
|--------|------|------|
| `Models/ProofreadModels.cs` | 复制 | 无 |
| `ILLMService.cs` | 复制 | 无 |
| `DeepSeekService.cs` | 复制 | 无 |
| `GLMService.cs` | 复制 | 无 |
| `OllamaService.cs` | 复制 | 无 |
| `LLMServiceFactory.cs` | 复制 | 无 |
| `HttpClientFactory.cs` | 复制 | 无 |
| `LLMRequestLogger.cs` | 复制 | Debug→Console/ILogger |
| `ProofreadIssueParser.cs` | 复制 | 无 |
| `DocumentSegmenter.cs` | 复制 | 无 |

### 5.2 需要改造

| 原文件 | 操作 | 改动详情 |
|--------|------|---------|
| `ProofreadService.cs` | 复制+改造 | ① `Dispatcher` → `IProgressReporter` 接口 ② 删除 `using System.Windows` ③ `Debug.WriteLine` → ILogger |
| `ProofreadCacheManager.cs` | 复制+改造 | 配置路径从 `%AppData%` 改为 `~/.config/gowordagent/` |
| `ConfigManager.cs` | 重写 | ① `ProtectedData` → AES 加密 ② 路径跨平台 ③ 保留数据模型 `AIConfig`/`ProviderConfig` 不变 |

### 5.3 需要新建

| 新文件 | 说明 |
|--------|------|
| `Program.cs` | ASP.NET Core 入口，配置 Kestrel、WebSocket、DI |
| `Controllers/ProofreadController.cs` | HTTP API 端点 |
| `Infrastructure/WsProgressReporter.cs` | WebSocket 进度推送 |
| `Infrastructure/CryptoService.cs` | AES 加密替代 DPAPI |
| `gowordagent-wps/plugin.json` | WPS 插件清单 |
| `gowordagent-wps/sidebar/index.html` | 侧边栏页面 |
| `gowordagent-wps/sidebar/js/*.js` | 前端逻辑（4 个文件） |
| `gowordagent-wps/sidebar/css/*.css` | 前端样式 |
| `install.sh` | 安装脚本 |
| `gowordagent.service` | systemd 用户服务 |

### 5.4 完全不碰

| 原文件 | 原因 |
|--------|------|
| `GOWordAgentPaneWpf.xaml.cs` | WPF 专属，新建 HTML 替代 |
| `GOWordAgentPaneHost.cs` | WPF 专属 |
| `WordDocumentService.cs` | Word COM 专属，新建 JS 替代 |
| `WordProofreadController.cs` | Word COM 专属，新建 JS 替代 |
| `ProofreadResultRenderer.cs` | WPF 专属 |
| `ViewModels/*.cs` | WPF MVVM 专属 |
| `gowordagentribbon.cs` | Word Ribbon 专属 |
| `ThisAddIn.cs` | VSTO 入口 |
| `*.xaml` | WPF 界面 |

---

## 六、工期估算（修正版）

### 6.1 按阶段拆解

| 阶段 | 工作项 | 工期 | 前置依赖 |
|------|--------|------|---------|
| **S0: POC** | WPS JS API 兼容性验证脚本 + 目标环境实测 | **3-5 天** | 需要麒麟/UOS + WPS 环境 |
| **S1: 后端骨架** | ASP.NET Core 项目 + ProofreadService 解耦 + CryptoService + HTTP API | **1 周** | 无 |
| **S2: 后端完善** | WebSocket 进度 + 配置 API + 连接测试 + 跨平台编译 | **1 周** | S1 |
| **S3: 前端 UI** | 侧边栏 HTML/CSS + 配置面板 + 消息列表 + 问题列表 | **1.5 周** | S2 |
| **S4: 前端文档操作** | document-service.js + proofread-controller.js + 逐 API 验证 | **1.5 周** | S0 + S3 |
| **S5: 集成调试** | 端到端联调 + 修复 WPS API 差异 + 安装脚本 | **1 周** | S2 + S4 |
| **S6: 测试** | 麒麟 x86 + UOS ARM64 全流程测试 | **1 周** | S5 |
| **合计** | | **6-8 周** | |

### 6.2 风险缓冲

S0（POC）如果发现关键 API 不可用，需要评估替代方案或降级策略，可能追加 1-2 周。

---

## 七、POC 验证步骤（建议第一周就做）

### Day 1-2：环境准备

```bash
# 1. 麒麟 V10 SP1 / UOS 21 安装 WPS Pro 11.8.2
# 2. 安装 .NET 8 Runtime
wget https://dot.net/v1/dotnet-install.sh
chmod +x dotnet-install.sh
./dotnet-install.sh --channel 8.0

# 3. 创建最简 WPS 插件
mkdir -p /tmp/gowordagent-poc/addons
```

### Day 2-3：运行兼容性测试

```bash
# 把 compat-test.js 代码放到 WPS 插件面板中
# 打开一个测试文档，逐项运行，记录结果
```

### Day 4-5：端到端验证

```bash
# 最简后端
dotnet new webapi -n gowordagent-poc-server
# 加一个 /api/proofread 端点，调用 DeepSeek API

# 最简前端
# 侧边栏：一个"校对"按钮，点后获取文档文本→调后端→拿到结果→显示

# 最简修订验证
# 拿到校对结果后，尝试用 WPS JS API 写修订
```

### POC 通过标准

- [ ] 能获取文档全文和选中文本
- [ ] 能调用后端 LLM API 并返回校对结果
- [ ] 能用修订模式写入文档（`TrackRevisions` + 替换文本）
- [ ] 能添加批注
- [ ] 能跳转定位
- [ ] WebSocket 进度推送正常
- [ ] 服务崩溃后能自动重启

**以上任何一项失败，都需要先找到替代方案再继续。**

---

## 八、备选方案与降级策略

| 场景 | 降级方案 |
|------|---------|
| `TrackRevisions` 不可用 | 直接替换文本 + 批注标记原文（现有代码已有降级逻辑，参考 `ApplyDegradedRevision`） |
| `Comments.Add` 不可用 | 不加批注，改为在侧边栏显示修改详情 |
| `Find.Execute` 行为不一致 | 改为纯 JS 字符串搜索 `Content.Text.indexOf()`，牺牲性能但保证正确 |
| `ScrollIntoView` 不可用 | 只 Select 不滚动，用户手动找 |
| WebSocket 不可靠 | 改为 HTTP 轮询（每 2 秒 GET 进度） |
| systemd 不可用 | shell 脚本 + crontab @reboot 启动 |

---

*本方案基于 gowordagent 源码实际分析，2026-03-31。建议先完成 S0 POC 验证再推进后续工作。*

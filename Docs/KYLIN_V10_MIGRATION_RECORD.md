# GOWordAgent 银河麒麟 V10 改造记录

> **项目**：gowordagent 银河麒麟 V10 适配改造  
> **版本**：v1.0  
> **日期**：2026-04-13  
> **改造目标**：将现有 Word VSTO 插件改造为 WPS Linux 版本

---

## 目录

1. [改造概述](#改造概述)
2. [项目结构调整](#项目结构调整)
3. [核心类库迁移 (GOWordAgent.Core)](#核心类库迁移)
4. [后端服务实现 (GOWordAgent.WpsService)](#后端服务实现)
5. [WPS 加载项开发 (GOWordAgent.WpsAddon)](#wps-加载项开发)
6. [配置管理适配 (跨平台加密)](#配置管理适配)
7. [API 设计实现](#api-设计实现)
8. [部署与安装](#部署与安装)
9. [问题与解决方案](#问题与解决方案)
10. [验证测试](#验证测试)

---

## 改造概述

### 原架构 vs 新架构

#### 原架构（Windows Word VSTO）
```
Word (Windows)
    ↓ VSTO Add-in
C# Plugin (GOWordAgentAddIn)
    - WPF UI (侧边栏)
    - Word COM Interop
    - LLM 服务调用
```

#### 新架构（银河麒麟 V10 + WPS）
```
银河麒麟 V10 桌面
│
├── WPS 文字 (Linux)
│   └── WPS 加载项 (HTML/JS/CSS)
│       ├── 设置面板
│       ├── 校对结果列表
│       └── HTTP 通信
│
└── .NET 8 后端服务 (Self-Contained)
    ├── Minimal API (Kestrel)
    ├── GOWordAgent.Core (复用核心逻辑)
    │   ├── ProofreadService
    │   ├── BaseLLMService
    │   ├── DocumentSegmenter
    │   └── Cache/Parser/Models
    └── 配置管理 (AES-GCM)
```

### 改造原则

1. **最大复用**：直接迁移现有 C# 核心逻辑，避免无意义的语言翻译
2. **前后分离**：WPF UI 改造为 HTML/JS，业务逻辑保留在 C# 后端
3. **跨平台**：.NET 8 Self-Contained 部署，零运行时依赖
4. **安全等价**：Windows DPAPI → Linux AES-GCM + /etc/machine-id

---

## 项目结构调整

### 目录结构对比

#### 改造前
```
GOWordAgentAddIn/
├── gowordagent.csproj          # VSTO 项目
├── gowordagent.sln
├── ThisAddIn.cs                # VSTO 入口
├── GOWordAgentPaneWpf.xaml     # WPF 侧边栏
├── GOWordAgentPaneWpf.xaml.cs  # 1200+ 行 UI 逻辑
├── WordDocumentService.cs      # Word COM 操作
├── WordProofreadController.cs  # 校对控制
├── ProofreadService.cs         # 校对服务核心
├── BaseLLMService.cs           # LLM 基类
├── ConfigManager.cs            # 配置管理 (DPAPI)
├── ...
```

#### 改造后
```
GOWordAgent/                              # 仓库根目录
│
├── GOWordAgentAddIn/                     # 【保留】现有 Word VSTO
│   ├── gowordagent.csproj
│   ├── GOWordAgentPaneWpf.xaml
│   ├── WordDocumentService.cs
│   └── ... (保持不变)
│
├── GOWordAgent.Core/                     # 【新增】.NET 8 共享类库
│   ├── GOWordAgent.Core.csproj
│   ├── Services/
│   │   ├── ProofreadService.cs           # 【迁移】校对服务
│   │   ├── BaseLLMService.cs             # 【迁移】LLM 基类
│   │   ├── DeepSeekService.cs            # 【迁移】DeepSeek 适配
│   │   ├── GLMService.cs                 # 【迁移】GLM 适配
│   │   ├── OllamaService.cs              # 【迁移】Ollama 适配
│   │   ├── DocumentSegmenter.cs          # 【迁移】文档分段
│   │   ├── ProofreadCacheManager.cs      # 【迁移】缓存管理
│   │   └── ProofreadIssueParser.cs       # 【迁移】结果解析
│   ├── Models/
│   │   └── ProofreadModels.cs            # 【迁移】数据模型
│   └── Config/
│       └── ConfigManager.cs              # 【适配】跨平台配置
│
├── GOWordAgent.WpsService/               # 【新增】.NET 8 Minimal API
│   ├── GOWordAgent.WpsService.csproj
│   ├── Program.cs                        # 服务入口
│   ├── appsettings.json
│   └── Controllers/
│       └── ProofreadController.cs        # API 控制器
│
├── GOWordAgent.WpsAddon/                 # 【新增】WPS 加载项
│   ├── package.json                      # wpsjs 配置
│   ├── index.html                        # 主页面
│   ├── main.js                           # 入口脚本
│   ├── ribbon.xml                        # 功能区配置
│   ├── css/
│   │   └── style.css                     # 样式
│   └── js/
│       ├── documentService.js            # WPS JS API 封装
│       ├── apiClient.js                  # 后端通信
│       ├── uiController.js               # UI 控制
│       └── proofreadService.js           # 校对工作流
│
└── Scripts/                              # 【新增】部署脚本
    ├── install.sh
    ├── uninstall.sh
    └── gowordagent.service
```

---

## 核心类库迁移

### 1. 项目创建

**文件**：`GOWordAgent.Core/GOWordAgent.Core.csproj`

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFramework>net8.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <AssemblyName>GOWordAgent.Core</AssemblyName>
    <RootNamespace>GOWordAgentAddIn</RootNamespace>
  </PropertyGroup>

  <ItemGroup>
    <!-- 保留 Newtonsoft.Json，不进行 System.Text.Json 迁移 -->
    <PackageReference Include="Newtonsoft.Json" Version="13.0.3" />
  </ItemGroup>

</Project>
```

**改造记录**：
- 使用 .NET 8 目标框架
- 保留 Newtonsoft.Json（文档明确要求，避免 AOT/Trimming 风险）
- 保持原有命名空间 `GOWordAgentAddIn`，减少代码改动

---

### 2. 直接迁移的文件（零逻辑改动）

| 文件 | 原位置 | 新位置 | 改动说明 |
|------|--------|--------|----------|
| ProofreadService.cs | GOWordAgentAddIn/ | GOWordAgent.Core/Services/ | 仅移除 WPF Dispatcher 依赖，改为事件回调 |
| BaseLLMService.cs | GOWordAgentAddIn/ | GOWordAgent.Core/Services/ | 零改动 |
| DeepSeekService.cs | GOWordAgentAddIn/ | GOWordAgent.Core/Services/ | 零改动 |
| GLMService.cs | GOWordAgentAddIn/ | GOWordAgent.Core/Services/ | 零改动 |
| OllamaService.cs | GOWordAgentAddIn/ | GOWordAgent.Core/Services/ | 零改动 |
| DocumentSegmenter.cs | GOWordAgentAddIn/ | GOWordAgent.Core/Services/ | 零改动 |
| ProofreadCacheManager.cs | GOWordAgentAddIn/ | GOWordAgent.Core/Services/ | 零改动 |
| ProofreadIssueParser.cs | GOWordAgentAddIn/ | GOWordAgent.Core/Services/ | 零改动 |
| Models/ProofreadModels.cs | GOWordAgentAddIn/Models/ | GOWordAgent.Core/Models/ | 零改动 |

**迁移代码示例**：

```csharp
// ProofreadService.cs 迁移改动
// 原代码（WPF 依赖）
private readonly Dispatcher _dispatcher;
await _dispatcher.InvokeAsync(() => OnProgress?.Invoke(this, args));

// 新代码（无 WPF 依赖）
public event EventHandler<ProofreadProgressArgs>? OnProgress;
OnProgress?.Invoke(this, args); // 直接触发事件，由消费者决定线程调度
```

---

## 后端服务实现

### 1. 项目创建

**文件**：`GOWordAgent.WpsService/GOWordAgent.WpsService.csproj`

```xml
<Project Sdk="Microsoft.NET.Sdk.Web">

  <PropertyGroup>
    <TargetFramework>net8.0</TargetFramework>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
    <AssemblyName>gowordagent-server</AssemblyName>
    
    <!-- Self-Contained 部署配置 -->
    <SelfContained>true</SelfContained>
    <RuntimeIdentifier>linux-x64</RuntimeIdentifier>
    <PublishSingleFile>true</PublishSingleFile>
    <PublishTrimmed>false</PublishTrimmed>
    <PublishReadyToRun>false</PublishReadyToRun>
  </PropertyGroup>

  <ItemGroup>
    <ProjectReference Include="..\GOWordAgent.Core\GOWordAgent.Core.csproj" />
  </ItemGroup>

</Project>
```

**配置说明**：
- `SelfContained=true`：生成独立部署包，无需目标机器安装 .NET Runtime
- `RuntimeIdentifier=linux-x64`：针对银河麒麟 x86_64 架构
- `PublishTrimmed=false`：禁用裁剪（方案文档明确要求，避免 Newtonsoft.Json 反射问题）

---

### 2. 服务入口 Program.cs

**文件**：`GOWordAgent.WpsService/Program.cs`

```csharp
using GOWordAgentAddIn;

var builder = WebApplication.CreateBuilder(args);

// 配置 Kestrel 监听地址（仅本地，安全）
builder.WebHost.ConfigureKestrel(options =>
{
    options.ListenLocalhost(0); // 0 = 自动分配端口
});

// 注册服务
builder.Services.AddSingleton<ProofreadService>();
builder.Services.AddSingleton<ILLMService>(sp => 
{
    // 从配置初始化默认 LLM 服务
    var config = ConfigManager.CurrentConfig;
    return LLMServiceFactory.CreateService(
        config.Provider, 
        config.ApiKey, 
        config.ApiUrl, 
        config.Model);
});

builder.Services.AddControllers();
builder.Services.AddCors(options =>
{
    options.AddPolicy("WpsAddon", policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyMethod()
              .AllowAnyHeader();
    });
});

var app = builder.Build();

app.UseCors("WpsAddon");
app.MapControllers();

// 获取实际分配的端口，写入端口文件
var port = app.Services.GetRequiredService<IHost>().Services
    .GetRequiredService<IServer>()
    .Features.Get<IEndpointFeature>()?
    .Endpoints.OfType<IPEndPoint>()
    .FirstOrDefault()?.Port ?? 0;

if (port > 0)
{
    var portFile = "/tmp/gowordagent-port.json";
    var portInfo = new
    {
        port,
        pid = Environment.ProcessId,
        timestamp = DateTimeOffset.UtcNow.ToUnixTimeSeconds()
    };
    File.WriteAllText(portFile, System.Text.Json.JsonSerializer.Serialize(portInfo));
    Console.WriteLine($"Service started on port {port}");
}

app.Run();

// 清理端口文件
if (File.Exists("/tmp/gowordagent-port.json"))
{
    try { File.Delete("/tmp/gowordagent-port.json"); } catch { }
}
```

---

### 3. API 控制器实现

**文件**：`GOWordAgent.WpsService/Controllers/ProofreadController.cs`

```csharp
using Microsoft.AspNetCore.Mvc;
using GOWordAgentAddIn;
using GOWordAgentAddIn.Models;

namespace GOWordAgent.WpsService.Controllers;

[ApiController]
[Route("api/[controller]")]
public class ProofreadController : ControllerBase
{
    private readonly ILLMService _llmService;
    private readonly ILogger<ProofreadController> _logger;

    public ProofreadController(ILLMService llmService, ILogger<ProofreadController> logger)
    {
        _llmService = llmService;
        _logger = logger;
    }

    /// <summary>
    /// 执行校对
    /// POST /api/proofread
    /// </summary>
    [HttpPost]
    public async Task<IActionResult> Proofread([FromBody] ProofreadRequest request)
    {
        try
        {
            // 创建校对服务
            var proofreadService = new ProofreadService(
                _llmService,
                request.Prompt,
                concurrency: 5,
                proofreadMode: request.Mode.ToString()
            );

            // 合并段落文本
            var fullText = string.Join("\n", request.Paragraphs.Select(p => p.Text));
            
            // 执行校对
            var results = await proofreadService.ProofreadDocumentAsync(fullText);
            
            // 转换为响应格式（包含偏移量）
            var response = ConvertToResponse(results, request.Paragraphs);
            
            return Ok(response);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Proofread failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>
    /// 健康检查
    /// GET /api/proofread/health
    /// </summary>
    [HttpGet("health")]
    public IActionResult Health()
    {
        return Ok(new { status = "ok", timestamp = DateTime.UtcNow });
    }

    private List<ProofreadResult> ConvertToResponse(
        List<ParagraphResult> results, 
        List<ParagraphInfo> paragraphs)
    {
        var response = new List<ProofreadResult>();
        
        foreach (var result in results)
        {
            var para = paragraphs[result.Index];
            var items = ProofreadIssueParser.ParseProofreadItems(result.ResultText);
            
            foreach (var item in items)
            {
                // 计算全局偏移量
                var startOffset = para.StartOffset + FindTextOffset(para.Text, item.Original);
                var endOffset = startOffset + item.Original.Length;
                
                response.Add(new ProofreadResult
                {
                    ParagraphIndex = result.Index,
                    StartOffset = startOffset,
                    EndOffset = endOffset,
                    Original = item.Original,
                    Suggestion = item.Modified,
                    Reason = item.Reason,
                    Severity = item.Severity,
                    Type = item.Type
                });
            }
        }
        
        return response;
    }

    private int FindTextOffset(string paragraph, string text)
    {
        var index = paragraph.IndexOf(text, StringComparison.Ordinal);
        return index >= 0 ? index : 0;
    }
}

// API 请求/响应模型
public class ProofreadRequest
{
    public string Text { get; set; } = "";
    public List<ParagraphInfo> Paragraphs { get; set; } = new();
    public string Provider { get; set; } = "DeepSeek";
    public string ApiKey { get; set; } = "";
    public string ApiUrl { get; set; } = "";
    public string Model { get; set; } = "";
    public string Prompt { get; set; } = "";
    public ProofreadMode Mode { get; set; } = ProofreadMode.Precise;
}

public class ParagraphInfo
{
    public int Index { get; set; }
    public int StartOffset { get; set; }
    public int EndOffset { get; set; }
    public string Text { get; set; } = "";
}

public class ProofreadResult
{
    public int ParagraphIndex { get; set; }
    public int StartOffset { get; set; }
    public int EndOffset { get; set; }
    public string Original { get; set; } = "";
    public string Suggestion { get; set; } = "";
    public string Reason { get; set; } = "";
    public string Severity { get; set; } = "";
    public string Type { get; set; } = "";
}

public enum ProofreadMode
{
    Precise,
    FullText
}
```

---

## WPS 加载项开发

### 1. 加载项配置

**文件**：`GOWordAgent.WpsAddon/package.json`

```json
{
  "name": "gowordagent-wps-addon",
  "version": "1.0.0",
  "description": "GOWordAgent WPS 加载项 - 银河麒麟 V10",
  "main": "main.js",
  "wps": {
    "id": "com.gowordagent.addin",
    "name": "智能校对",
    "version": "1.0.0",
    "host": "wpp,wet", 
    "url": "index.html"
  }
}
```

---

### 2. 主页面 HTML

**文件**：`GOWordAgent.WpsAddon/index.html`

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>智能校对</title>
    <link rel="stylesheet" href="css/style.css">
</head>
<body>
    <div id="app">
        <!-- 状态栏 -->
        <div id="status-bar" class="status-bar">
            <span id="connection-status">未连接</span>
        </div>
        
        <!-- 设置面板 -->
        <div id="settings-panel" class="panel">
            <h3>AI 配置</h3>
            <div class="form-group">
                <label>AI 提供商</label>
                <select id="provider-select">
                    <option value="DeepSeek">DeepSeek</option>
                    <option value="GLM">智谱 AI</option>
                    <option value="Ollama">本地 Ollama</option>
                </select>
            </div>
            <div class="form-group">
                <label>API Key</label>
                <input type="password" id="api-key" placeholder="输入 API Key">
            </div>
            <div class="form-group">
                <label>模型</label>
                <input type="text" id="model" placeholder="deepseek-chat">
            </div>
            <div class="form-group">
                <label>校验模式</label>
                <div class="radio-group">
                    <label><input type="radio" name="mode" value="Precise" checked> 精准校验</label>
                    <label><input type="radio" name="mode" value="FullText"> 全文校验</label>
                </div>
            </div>
            <button id="btn-connect" class="btn-primary">保存并连接</button>
        </div>
        
        <!-- 结果面板 -->
        <div id="results-panel" class="panel hidden">
            <h3>校对结果</h3>
            <div id="progress-info"></div>
            <div id="issues-list"></div>
        </div>
        
        <!-- 操作栏 -->
        <div class="action-bar">
            <button id="btn-proofread" class="btn-primary" disabled>开始校对</button>
        </div>
    </div>
    
    <script src="js/apiClient.js"></script>
    <script src="js/documentService.js"></script>
    <script src="js/proofreadService.js"></script>
    <script src="js/uiController.js"></script>
    <script src="main.js"></script>
</body>
</html>
```

---

### 3. WPS JS API 封装

**文件**：`GOWordAgent.WpsAddon/js/documentService.js`

```javascript
/**
 * WPS 文档操作服务
 * 封装 WPS JS API，提供文档读写接口
 */

var DocumentService = {
    /**
     * 获取文档文本（带偏移量信息）
     * @returns {Array} 段落数组，包含索引、文本、偏移量
     */
    getDocumentText: function() {
        var app = wps.WpsApplication();
        var doc = app.ActiveDocument;
        
        if (!doc) {
            throw new Error("没有打开的文档");
        }
        
        var paragraphs = [];
        var offset = 0;
        
        for (var i = 1; i <= doc.Paragraphs.Count; i++) {
            var p = doc.Paragraphs.Item(i);
            var text = p.Range.Text;
            
            paragraphs.push({
                index: i - 1,  // 0-based index
                start: offset,
                end: offset + text.length,
                text: text
            });
            
            offset += text.length;
        }
        
        return paragraphs;
    },
    
    /**
     * 在指定偏移量位置应用修订
     * @param {number} startOffset - 起始偏移量
     * @param {number} endOffset - 结束偏移量
     * @param {string} replacement - 替换文本
     * @param {string} comment - 批注内容
     */
    applyAtOffset: function(startOffset, endOffset, replacement, comment) {
        var app = wps.WpsApplication();
        var doc = app.ActiveDocument;
        
        // 定位到指定范围
        var range = doc.Range(startOffset, endOffset);
        
        // 开启修订模式
        var oldTrackRevisions = doc.TrackRevisions;
        doc.TrackRevisions = true;
        
        try {
            // 删除原文并插入新文本
            range.Delete();
            range.InsertAfter(replacement);
            
            // 添加批注
            if (comment) {
                doc.Comments.Add(range, comment);
            }
        } finally {
            doc.TrackRevisions = oldTrackRevisions;
        }
    },
    
    /**
     * 导航到指定偏移量位置
     * @param {number} startOffset - 起始偏移量
     * @param {number} endOffset - 结束偏移量
     */
    navigateToOffset: function(startOffset, endOffset) {
        var app = wps.WpsApplication();
        var doc = app.ActiveDocument;
        
        var range = doc.Range(startOffset, endOffset);
        range.Select();
        
        // 滚动到视图中
        var window = app.ActiveWindow;
        window.ScrollIntoView(range);
    },
    
    /**
     * 获取选中的文本
     * @returns {string}
     */
    getSelectedText: function() {
        var app = wps.WpsApplication();
        var selection = app.Selection;
        return selection ? selection.Text : "";
    }
};
```

---

### 4. 后端通信客户端

**文件**：`GOWordAgent.WpsAddon/js/apiClient.js`

```javascript
/**
 * 后端 API 通信客户端
 * 使用 XMLHttpRequest（兼容性最广）
 */

var ApiClient = {
    baseUrl: '',
    
    /**
     * 发现后端服务端口
     */
    discoverService: function() {
        try {
            var portData = wps.FileSystem.ReadFile('/tmp/gowordagent-port.json');
            var info = JSON.parse(portData);
            this.baseUrl = 'http://127.0.0.1:' + info.port;
            return true;
        } catch (e) {
            console.error('Service discovery failed:', e);
            return false;
        }
    },
    
    /**
     * GET 请求
     */
    get: function(path, callback) {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', this.baseUrl + path, true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        
        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4) {
                if (xhr.status === 200) {
                    callback(null, JSON.parse(xhr.responseText));
                } else {
                    callback(new Error('HTTP ' + xhr.status), null);
                }
            }
        };
        
        xhr.send();
    },
    
    /**
     * POST 请求
     */
    post: function(path, data, callback) {
        var xhr = new XMLHttpRequest();
        xhr.open('POST', this.baseUrl + path, true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        
        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4) {
                if (xhr.status === 200) {
                    callback(null, JSON.parse(xhr.responseText));
                } else {
                    callback(new Error('HTTP ' + xhr.status), null);
                }
            }
        };
        
        xhr.send(JSON.stringify(data));
    },
    
    /**
     * 健康检查
     */
    healthCheck: function(callback) {
        this.get('/api/proofread/health', callback);
    },
    
    /**
     * 执行校对
     */
    proofread: function(data, callback) {
        this.post('/api/proofread', data, callback);
    }
};
```

---

## 配置管理适配

### 跨平台加密实现

**文件**：`GOWordAgent.Core/Config/ConfigManager.cs`

```csharp
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// 跨平台配置管理器
    /// Windows: DPAPI + HMAC
    /// Linux:   AES-GCM + /etc/machine-id
    /// </summary>
    public static class ConfigManager
    {
        // Linux 配置路径
        private static readonly string LinuxConfigDir = 
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".config", "gowordagent");
        private static readonly string LinuxConfigFile = 
            Path.Combine(LinuxConfigDir, "config.dat");
        
        // Windows 配置路径（保持兼容）
        private static readonly string WindowsConfigDir = 
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
            "GOWordAgentAddIn");
        private static readonly string WindowsConfigFile = 
            Path.Combine(WindowsConfigDir, "config.dat");
        
        private static string ConfigFile => 
            RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? WindowsConfigFile : LinuxConfigFile;
        
        public static AIConfig CurrentConfig { get; private set; } = new AIConfig();
        
        /// <summary>
        /// 加载配置
        /// </summary>
        public static void LoadConfig()
        {
            try
            {
                if (!File.Exists(ConfigFile))
                {
                    CurrentConfig = new AIConfig();
                    return;
                }
                
                byte[] encrypted = File.ReadAllBytes(ConfigFile);
                byte[] decrypted;
                
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    // Windows: 使用 DPAPI
                    decrypted = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
                }
                else
                {
                    // Linux: 使用 AES-GCM
                    decrypted = LinuxCrypto.Decrypt(encrypted);
                }
                
                string json = Encoding.UTF8.GetString(decrypted);
                var config = JsonConvert.DeserializeObject<AIConfig>(json);
                CurrentConfig = config ?? new AIConfig();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Load config failed: {ex.Message}");
                CurrentConfig = new AIConfig();
            }
        }
        
        /// <summary>
        /// 保存配置
        /// </summary>
        public static void SaveConfig(AIConfig config)
        {
            try
            {
                string json = JsonConvert.SerializeObject(config, Formatting.Indented);
                byte[] plainBytes = Encoding.UTF8.GetBytes(json);
                byte[] encrypted;
                
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    encrypted = ProtectedData.Protect(plainBytes, null, DataProtectionScope.CurrentUser);
                }
                else
                {
                    encrypted = LinuxCrypto.Encrypt(plainBytes);
                }
                
                Directory.CreateDirectory(Path.GetDirectoryName(ConfigFile)!);
                File.WriteAllBytes(ConfigFile, encrypted);
                CurrentConfig = config;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Save config failed: {ex.Message}");
                throw;
            }
        }
        
        // ... 其他方法保持不变
    }
    
    /// <summary>
    /// Linux 加密实现 (AES-GCM + /etc/machine-id)
    /// </summary>
    public static class LinuxCrypto
    {
        private static byte[] DeriveKey()
        {
            var machineId = File.ReadAllText("/etc/machine-id").Trim();
            return SHA256.HashData(Encoding.UTF8.GetBytes(machineId));
        }
        
        public static byte[] Encrypt(byte[] plainData)
        {
            var key = DeriveKey();
            var nonce = RandomNumberGenerator.GetBytes(12);
            
            using var aes = new AesGcm(key, 16);
            var cipherData = new byte[plainData.Length];
            var tag = new byte[16];
            
            aes.Encrypt(nonce, plainData, cipherData, tag);
            
            // 组合: nonce(12) + tag(16) + ciphertext
            var result = new byte[12 + 16 + cipherData.Length];
            Buffer.BlockCopy(nonce, 0, result, 0, 12);
            Buffer.BlockCopy(tag, 0, result, 12, 16);
            Buffer.BlockCopy(cipherData, 0, result, 28, cipherData.Length);
            return result;
        }
        
        public static byte[] Decrypt(byte[] encryptedData)
        {
            var key = DeriveKey();
            var nonce = new byte[12];
            var tag = new byte[16];
            var cipherData = new byte[encryptedData.Length - 28];
            
            Buffer.BlockCopy(encryptedData, 0, nonce, 0, 12);
            Buffer.BlockCopy(encryptedData, 12, tag, 0, 16);
            Buffer.BlockCopy(encryptedData, 28, cipherData, 0, cipherData.Length);
            
            using var aes = new AesGcm(key, 16);
            var plainData = new byte[cipherData.Length];
            aes.Decrypt(nonce, cipherData, tag, plainData);
            return plainData;
        }
    }
}
```

**改造要点**：
- 使用 `RuntimeInformation.IsOSPlatform` 判断操作系统
- Windows 保持原有 DPAPI 实现
- Linux 使用 AES-GCM + /etc/machine-id 派生密钥
- 配置文件路径遵循 XDG 规范（~/.config/gowordagent/）

---

## 部署与安装

### 1. Systemd 用户服务

**文件**：`Scripts/gowordagent.service`

```ini
[Unit]
Description=GOWordAgent Backend Service
After=network.target

[Service]
Type=simple
ExecStart=/opt/gowordagent/gowordagent-server
Restart=on-failure
RestartSec=3
StandardOutput=journal
StandardError=journal

[Install]
WantedBy=default.target
```

### 2. 安装脚本

**文件**：`Scripts/install.sh`

```bash
#!/bin/bash
set -e

INSTALL_DIR=/opt/gowordagent
CONFIG_DIR=$HOME/.config/gowordagent
ADDON_DIR=$HOME/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin

echo "=== GOWordAgent 安装脚本 ==="

# 1. 检查架构
ARCH=$(uname -m)
if [ "$ARCH" != "x86_64" ]; then
    echo "错误: 当前仅支持 x86_64 架构，检测到 $ARCH"
    exit 1
fi

# 2. 检查 WPS 版本
if ! command -v wps &> /dev/null; then
    echo "错误: 未检测到 WPS Office"
    exit 1
fi

# 3. 复制后端
echo "正在安装后端服务..."
sudo mkdir -p $INSTALL_DIR
sudo cp -r ./backend/* $INSTALL_DIR/
sudo chmod +x $INSTALL_DIR/gowordagent-server

# 4. 注册 systemd 用户服务
mkdir -p $HOME/.config/systemd/user
cp ./scripts/gowordagent.service $HOME/.config/systemd/user/
systemctl --user daemon-reload
systemctl --user enable gowordagent
systemctl --user start gowordagent

# 5. 安装 WPS 加载项
echo "正在安装 WPS 加载项..."
mkdir -p $ADDON_DIR
cp -r ./addon/* $ADDON_DIR/

# 6. 创建配置目录
mkdir -p $CONFIG_DIR

echo "=== 安装完成 ==="
echo "后端服务状态:"
systemctl --user status gowordagent --no-pager
echo ""
echo "请重启 WPS 文字以加载插件"
```

### 3. 卸载脚本

**文件**：`Scripts/uninstall.sh`

```bash
#!/bin/bash

echo "=== GOWordAgent 卸载脚本 ==="

# 停止并禁用服务
systemctl --user stop gowordagent 2>/dev/null || true
systemctl --user disable gowordagent 2>/dev/null || true
rm -f $HOME/.config/systemd/user/gowordagent.service
systemctl --user daemon-reload

# 删除文件
sudo rm -rf /opt/gowordagent
rm -rf $HOME/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin

echo "=== 卸载完成 ==="
echo "配置文件保留在: $HOME/.config/gowordagent/"
echo "如需完全清理，请手动删除该目录"
```

---

## 问题与解决方案

### 问题 1：WPF Dispatcher 依赖

**现象**：ProofreadService 原代码依赖 WPF Dispatcher 进行线程调度

**解决方案**：
- 移除 Dispatcher 依赖
- 改为纯事件驱动，由消费者决定线程调度
- 后端服务中直接使用 Task，无 UI 线程要求

### 问题 2：HTTP 通信跨域

**现象**：WPS 加载项 WebView 与本地后端通信可能遇到 CORS 限制

**解决方案**：
- 后端启用 CORS，允许任意来源（仅本地）
- WPS 加载项使用 `127.0.0.1` 而非 `localhost`
- 准备长轮询降级方案（若 SSE 不支持）

### 问题 3：文本定位精度

**现象**：WPS JS API 的 Find.Execute 可能匹配错误位置

**解决方案**：
- 采用字符偏移量定位策略
- 后端返回每个校对项的 StartOffset/EndOffset
- 前端使用 Range(start, end) 精确定位

### 问题 4：配置加密跨平台

**现象**：Windows DPAPI 在 Linux 上不可用

**解决方案**：
- Linux 使用 AES-GCM 对称加密
- 密钥从 /etc/machine-id 派生
- 保持加密强度与 DPAPI 相当

---

## 验证测试

### 功能测试清单

| 测试项 | 期望结果 | 状态 |
|--------|----------|------|
| 后端服务启动 | 端口文件写入 /tmp/gowordagent-port.json | 待测试 |
| 健康检查 API | GET /api/proofread/health 返回 ok | 待测试 |
| 校对 API | POST /api/proofread 返回校对结果 | 待测试 |
| WPS 加载项加载 | 侧边栏显示智能校对面板 | 待测试 |
| 连接后端 | 显示已连接状态 | 待测试 |
| 提取文档文本 | 正确获取段落内容和偏移量 | 待测试 |
| 执行校对 | 返回校对结果列表 | 待测试 |
| 写入修订 | 文档显示修订和批注 | 待测试 |
| 点击定位 | 跳转到文档对应位置 | 待测试 |
| 配置保存 | 重启后配置保留 | 待测试 |

### 兼容性测试

| 环境 | 版本 | 状态 |
|------|------|------|
| 银河麒麟 V10 SP1 | - | 待测试 |
| WPS Office for Linux | 12.1.2.25838 | 待测试 |
| .NET 8 Runtime | 8.0.x | 无需安装（Self-Contained） |

---

## 总结

### 改造完成项

1. **GOWordAgent.Core** - .NET 8 共享类库，迁移所有核心逻辑
2. **GOWordAgent.WpsService** - Minimal API 后端服务
3. **GOWordAgent.WpsAddon** - WPS 加载项前端
4. **跨平台 ConfigManager** - AES-GCM 加密适配 Linux
5. **部署脚本** - Systemd 服务 + 安装/卸载脚本

### 关键设计决策

| 决策 | 选择 | 理由 |
|------|------|------|
| 前端框架 | 原生 HTML/JS | 避免构建链复杂性和 WebView 兼容性问题 |
| JSON 库 | Newtonsoft.Json | 现有代码稳定，避免迁移风险 |
| 部署方式 | Self-Contained | 零运行时依赖，信创桌面可接受 |
| 加密方案 | AES-GCM + machine-id | Linux 原生兼容，无需额外库 |
| 文本定位 | 偏移量策略 | 避免 Find.Execute 重复匹配风险 |

### 下一步工作

1. Day 1 PoC：验证 WPS 加载项与本地 HTTP 通信
2. Day 2-5：后端 + UI 开发调试
3. Day 6-8：WPS JS API 集成
4. Day 9-10：麒麟 x86_64 部署验证

---

*文档结束*

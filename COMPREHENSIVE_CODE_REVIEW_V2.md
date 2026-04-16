# GOWordAgent 全面代码审查报告 V2（Linux 视角）

**审查日期**: 2026-04-14  
**审查范围**: 全部源码（含上轮修复后增量）  
**审查维度**: 功能完成度、架构设计、代码质量、健壮性、安全性、性能、Linux 兼容性  
**构建状态**: ✅ `dotnet build GOWordAgent.sln` — 0 警告 0 错误

---

## ⚠️ 致命缺陷（必须立即修复）

### 1. `ProofreadService.CancelAll()` 具有“永久性杀伤”效果
**文件**: `GOWordAgent.Core/Services/ProofreadService.cs`

**问题描述**:
```csharp
private static readonly CancellationTokenSource _globalCancelCts = new CancellationTokenSource();
```
`CancelAll()` 调用 `_globalCancelCts.Cancel()` 后，该静态 `CTS` **永远处于 `IsCancellationRequested = true` 状态**。而 `ResetGlobalCancellation()` 是**空实现**，没有任何替换逻辑。

**致命后果**:
- 用户或前端调用一次 `POST /api/proofread/cancel` 后，所有**后续**的校对请求在创建 `LinkedTokenSource` 时会立即继承已取消状态，导致 `ProofreadDocumentAsync` 一启动就抛出 `OperationCanceledException`。
- **校对功能永久瘫痪**，必须重启后端进程才能恢复。

**修复建议**:
```csharp
private static CancellationTokenSource _globalCancelCts = new CancellationTokenSource();

public static void CancelAll()
{
    lock (_globalCancelLock)
    {
        if (!_globalCancelCts.IsCancellationRequested)
        {
            try { _globalCancelCts.Cancel(); }
            catch (ObjectDisposedException) { }
        }
    }
}

public static void ResetGlobalCancellation()
{
    lock (_globalCancelLock)
    {
        var old = _globalCancelCts;
        _globalCancelCts = new CancellationTokenSource();
        try { old.Dispose(); }
        catch { }
    }
}
```
然后在 `ProofreadController.Proofread()` 的 `catch (OperationCanceledException)` 分支末尾调用 `ProofreadService.ResetGlobalCancellation()`，确保取消状态被重置。或者更优雅地：每次进入 `Proofread` Action 时先 `ResetGlobalCancellation()`。

---

## 🔴 高优先级问题

### 2. `LLMServiceFactory` 对 Ollama 的参数传递错误
**文件**: `GOWordAgent.Core/Services/LLMServiceFactory.cs`

```csharp
case AIProvider.Ollama:
    return new OllamaService(apiUrl ?? "http://localhost:11434", apiKey);
```

而 `OllamaService` 构造函数签名：
```csharp
public OllamaService(string apiUrl, string? model = null)
```

**后果**: 用户输入的 `apiKey` 被错误地当作 `model` 名称传入 Ollama。由于 Ollama 本地服务不需要 API Key，但如果用户在 UI 填写了任意内容（包括空字符串），它会被当成模型名，导致 Ollama 返回 `model not found`。

**修复建议**:
```csharp
case AIProvider.Ollama:
    return new OllamaService(apiUrl ?? "http://localhost:11434", model);
```

### 3. `ProofreadCacheManager` 内存计数器在 `TryAdd` 失败时仍累加
**文件**: `GOWordAgent.Core/Services/ProofreadCacheManager.cs`

```csharp
if (_globalCache.TryAdd(cacheKey, result))
{
    _accessCount[cacheKey] = Interlocked.Increment(ref _accessCounter);
    Interlocked.Add(ref _estimatedBytes, itemBytes);
}
```

这段代码看起来正确，但问题在 `StoreResult` 入口的 `while` 循环：
```csharp
while (Interlocked.Read(ref _estimatedBytes) + itemBytes > MaxCacheBytes && !_globalCache.IsEmpty)
{
    EvictOldestEntries(10);
}
```

高并发下，多个线程可能同时看到 `_estimatedBytes + itemBytes > MaxCacheBytes`，然后都执行淘汰，导致过度淘汰。更严重的是，如果 `cacheKey` 已经存在（例如同一文本被并发处理），`TryAdd` 失败，但 `_estimatedBytes` **没有增加**（当前代码是正确的），这本身没问题。

然而，`EvictOldestEntries` 中通过 `EstimateResultBytes(removed)` 来扣减 `_estimatedBytes`。如果 `removed` 为 null（`TryRemove` 失败），则不扣减。这在并发下可能导致 `_estimatedBytes` 被**重复扣减**（多个线程同时淘汰同一 key），造成计数器偏低甚至负数。

**修复建议**: 在 `EvictOldestEntries` 中，先收集 key，再逐个 `TryRemove`，但把扣减逻辑放在 `TryRemove` 成功后才执行（当前已经是这样）。真正的问题是多个线程可能同时读取 `_estimatedBytes` 并进入 `while` 循环，导致重复淘汰。建议把 `while` 条件改为基于 `TryAdd` 后的状态来懒淘汰，或者使用 `Interlocked.CompareExchange` 风格的循环来限制。

更简单的修复：在 `while` 循环内加一次 `Interlocked.Read`，并在 `TryAdd` 后如果失败不调整 `_estimatedBytes`（当前已满足）。主要修复过度淘汰即可接受。

### 4. 前端缺少“取消校对”按钮
虽然后端提供了 `POST /api/proofread/cancel`，但 `index.html` 中没有任何 UI 元素调用 `ApiClient.cancel()`。用户在校对过程中（尤其是长文档）无法主动中断，必须关闭 WPS 或等待超时。

**修复建议**: 在 `index.html` 的操作栏增加 `<button id="btn-cancel">取消</button>`，并在 `proofreadService.js` 的 `startProofread` 中绑定取消逻辑。

### 5. `ApiClient.rediscover()` 在 401 后重试存在逻辑漏洞
```javascript
if (self.rediscover()) {
    self.get(path, callback, true);
}
```

`rediscover()` 只是重新读取端口文件并设置 `baseUrl`。如果旧的后端进程已经退出且新进程尚未启动（端口文件不存在），`rediscover()` 返回 `false`，请求直接失败。但如果旧进程还在（端口文件还在），`rediscover()` 返回 `true`，但此时它读取的仍是**旧 Token**，因为旧进程没更新端口文件。这意味着 401 后会无限重试（虽然 `_retried` 限制了只重试一次，但用户体验仍是直接报错）。

**更深层问题**: 401 通常意味着后端重启生成了新 Token。如果新进程的端口文件尚未写入（启动延迟），前端在 401 后重试仍然会失败。这不是代码 bug，而是用户体验问题。建议在 UI 上提示“服务已重启，请重新保存配置”而不是静默重试。

---

## 🟡 中优先级问题

### 6. `ProofreadController.Cancel()` 未限流且未重置取消状态
`POST /api/proofread/cancel` 没有应用 `[EnableRateLimiting]`。虽然本地服务场景下影响有限，但恶意/误触的连续调用会不断调用 `CancelAll()`。

此外，如致命缺陷#1所述，该端点调用后**没有重置**取消状态，导致服务瘫痪。

### 7. 部署脚本 `probe_wps_addon_dirs` 探测逻辑不准确
**文件**: `deploy-linux.sh`, `Scripts/install.sh`

```bash
for d in "${dirs[@]}"; do
    if [ -d "$(dirname "$d")" ]; then
        echo "$d"
        return
    fi
done
```

该逻辑检查的是**加载项目录的父目录**（如 `jsaddons`）是否存在。如果 WPS 安装后从未加载过任何插件，`jsaddons` 目录可能不存在，脚本就会 fallback 到默认路径，而实际上 WPS 可执行文件安装在 `/opt/kingsoft/wps-office/`，正确的加载项目录也应在该路径下。这导致在某些干净系统上插件安装到了错误位置。

**修复建议**: 探测 WPS 安装根目录（通过 `which wps` 或常见安装路径），然后推导 `jsaddons` 路径，而不是反向检查 `jsaddons` 父目录是否存在。

### 8. 并发度固定为 5，未区分提供商
**文件**: `ProofreadController.cs`

```csharp
using var proofreadService = new ProofreadService(
    llmService,
    prompt,
    concurrency: 5,
    proofreadMode: request.Mode.ToString()
);
```

Ollama 本地推理对并发敏感，5 并发可能直接打满本地 GPU/CPU 显存或内存，导致所有请求超时。DeepSeek/GLM 在线 API 则更适合 5 并发。

**修复建议**: 在 `LLMServiceFactory` 或各服务上增加 `RecommendedConcurrency` 属性，`Ollama` 默认 1~2，在线 API 默认 5。

### 9. `ConvertToResponse` 仍对失败段落进行偏移量计算过滤
当前 `ConvertToResponse` 通过 `if (!result.IsCompleted) continue;` 跳过失败段落，这是正确的。但 `result.Items` 在失败时是一个空列表，所以不会影响。这里没有问题，只是备注说明已正确处理。

### 10. 线程池设置可能过于激进
**文件**: `Program.cs`

```csharp
ThreadPool.SetMaxThreads(Math.Max(32, processorCount * 8), Math.Max(32, processorCount * 8));
```

在 64 核服务器上，这将设置 512 个线程。虽然是本地服务，但 `Parallel.ForEachAsync` + HttpClient 的并发模型并不依赖如此高的线程上限，反而可能在负载突增时导致内存飙升。建议上限设为 `Math.Min(128, Math.Max(32, processorCount * 4))`。

---

## 🟢 低优先级问题

### 11. `ProofreadService.ResetGlobalCancellation()` 空实现误导维护者
当前该方法包含大量注释但没有任何实际代码，容易让后续开发者误以为取消状态已被重置。建议要么实现，要么删除该方法（如果已经在其他地方处理）。鉴于致命缺陷#1，必须实现。

### 12. 前端 `manualPort` 输入缺少格式校验
`UIController.getConfig()` 中：
```javascript
var manualPort = parseInt(manualPortStr, 10);
```

如果用户输入 `abc`，`parseInt` 返回 `NaN`，当前代码会回退到 0，但没有 UI 提示。建议增加 `isNaN` 时的输入框红色边框提示。

### 13. `/api/proofread/stats` 和 `/api/proofread/clear-cache` 的权限
`GetStats` 没有应用 `[EnableRateLimiting]`，但这两个端点都是只读/管理操作，影响较小。`ClearCache` 已有限流，合理。

### 14. `ProofreadController.cs` 中的 `Catch` 块缺少 `ResetGlobalCancellation`
`Proofread` Action 的 `catch (OperationCanceledException)` 和 `catch (Exception)` 中，如果异常是由 `_globalCancelCts` 触发的，后续请求仍会继承该取消状态。应在所有退出路径上重置，或在 Action 入口处统一重置。

### 15. `ApiTokenAuth.Validate` 仍对 `/config` 免 Token
这是兼容旧版前端的必要妥协，但随着前端已经支持 `rediscover()` 和 `apiToken` 持久化，建议在 1~2 个版本后移除 `/config` 的豁免，增强安全性。

---

## 各维度评分（修复前）

| 维度 | 评分 | 关键扣分点 |
|------|------|------------|
| 功能完成度 | ⭐⭐⭐☆☆ | **CancelAll 永久瘫痪后续请求**、Ollama 参数传错、缺少取消按钮 |
| 架构设计 | ⭐⭐⭐⭐☆ | 静态全局 CTS 管理不当，其余分层良好 |
| 代码质量 | ⭐⭐⭐⭐☆ | 存在参数传递低级错误，其余规范 |
| 健壮性 | ⭐⭐⭐☆☆ | 全局取消状态无法恢复是重大可用性灾难 |
| 安全性 | ⭐⭐⭐⭐☆ | CORS 已修复，chmod 已加固，主要风险在 `/config` 免 Token |
| 性能 | ⭐⭐⭐⭐☆ | 缓存统计已优化为 O(1)，但并发度未按提供商区分 |
| Linux 兼容性 | ⭐⭐⭐⭐☆ | WPS 路径探测、手动端口兜底、脚本改进显著，但 `probe_wps_addon_dirs` 仍不够精确 |

**综合评分**: ⭐⭐⭐☆☆ (5.5/10)

> 注：若修复致命缺陷 #1 和高优先级问题 #2，评分可迅速回升至 8.0/10。

---

## 修复优先级清单

1. **立即修复 `ProofreadService` 全局 CTS 不可重置问题**（致命）。
2. **修复 `LLMServiceFactory` 中 Ollama 的 `apiKey` 误传为 `model`**（高）。
3. **前端增加“取消校对”按钮并调用 `ApiClient.cancel()`**（高）。
4. **在 `ProofreadController.Proofread` Action 入口处调用 `ResetGlobalCancellation()`**（高，作为 #1 的配套修复）。
5. **优化 `ProofreadCacheManager` 的并发淘汰逻辑**，防止重复扣减（中）。
6. **改进 `probe_wps_addon_dirs` 的探测逻辑**，基于 WPS 安装根目录推导（中）。
7. **为不同 LLM 提供商设置不同的默认并发度**（中）。
8. **收紧线程池最大线程上限**（低）。

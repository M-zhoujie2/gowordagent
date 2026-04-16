# GOWordAgent 全面代码审查报告（Linux 视角）

**审查日期**: 2026-04-14  
**审查范围**: `GOWordAgent.Core` / `GOWordAgent.WpsService` / `GOWordAgent.WpsAddon` / 部署脚本  
**审查维度**: 功能完成度、架构设计、代码质量、健壮性、安全性、性能、Linux 兼容性  
**构建状态**: ✅ `dotnet build GOWordAgent.sln` — 0 警告 0 错误

---

## 1. 功能完成度

### 现状
项目已形成完整的端到端闭环：WPS 加载项提取文档段落 → ASP.NET Core 本地服务 → 并发 LLM 校对 → 缓存 → 解析结果 → 回写 WPS 修订/批注。支持 DeepSeek、GLM、Ollama 三种提供商，具备配置持久化、服务发现、systemd 托管等能力。

### 问题与风险

| 优先级 | 问题 | 影响 | 建议 |
|--------|------|------|------|
| **高** | **缺少显式取消端点/机制** | 用户点击 UI "取消" 后，后端 `Parallel.ForEachAsync` 仍在继续消耗 Token 和 CPU；控制器仅靠 `HttpContext.RequestAborted` 被动取消，若连接未断开则无法中断。 | 新增 `POST /api/proofread/cancel` 或 `CancellationTokenSource` 注册表，支持服务端主动中断。 |
| **高** | **进度事件未反馈到前端** | `ProofreadService.OnProgress` 在控制器中无人订阅，前端只能干等最终响应，用户体验差。 | 使用 SSE（Server-Sent Events）或拆分为轮询进度 API，将节流后的进度实时推送到前端。 |
| **中** | **失败时丢失所有部分结果** | `Parallel.ForEachAsync` 遇到单段 LLM 异常后全部取消，前面 N-1 段已成功校对的成果被丢弃。 | 对每段执行 `try/catch` 并记录失败段落，最终返回 `{ completedIssues, failedParagraphIndices }`，让用户至少能看到已处理的部分。 |
| **中** | **后端重启后前端 Token 失效** | `ApiTokenAuth.Token` 每次启动重新生成，前端缓存的旧 Token 会导致 401，且 UI 没有自动重发现逻辑。 | 在 `ApiClient` 捕获 401 后自动触发 `discoverService()` 重新读取端口文件并更新 Token。 |
| **低** | **缺少缓存/指标管理端点** | 无法在线查看缓存命中率、内存占用，也无法远程清空缓存。 | 增加 `/api/proofread/stats` 和 `/api/proofread/clear-cache` 端点（受 Token 保护）。 |

---

## 2. 架构设计

### 优点
- **三层分离清晰**: Core（纯业务逻辑）/ WpsService（ASP.NET 宿主）/ WpsAddon（前端交互）。
- **工厂模式**: `LLMServiceFactory` + `BaseLLMService` 抽象使得新增提供商成本极低。
- **无状态控制器**: 每个请求独立创建服务实例，天然线程安全。

### 问题与风险

| 优先级 | 问题 | 说明 | 建议 |
|--------|------|------|------|
| **中** | **全局静态缓存难以测试和替换** | `ProofreadCacheManager` 是纯静态类，隐藏全局状态，单元测试时无法 mock 或隔离。 | 抽象出 `IProofreadCache` 接口，在 `ProofreadService` 中通过构造函数注入，默认实现仍用 `ConcurrentDictionary`。 |
| **中** | **控制器与进度报告职责耦合** | `ProofreadService` 既负责 LLM 编排，又负责 UI 节流进度报告。 | 将进度报告提取为 `IProgressReporter` 接口，控制器注入 `NullProgressReporter` 或 `HttpProgressReporter`。 |
| **低** | **Core 层完全无 DI** | `ConfigManager`、`ProofreadCacheManager` 均为静态工具类，无法通过 IoC 容器管理生命周期。 | 对于当前规模可接受；若后续扩展，建议将 Core 服务注册为 `Scoped`/`Singleton`。 |
| **低** | **JSON 库混用** | Core 和控制器使用 `Newtonsoft.Json`，而 `Program.cs` 写端口文件时使用了 `System.Text.Json`。 | 统一使用 `System.Text.Json`（.NET 8 性能更好），或至少在 `Program.cs` 中也使用 `Newtonsoft.Json` 保持一致。 |

---

## 3. 代码质量

### 优点
- 输入校验充分（`BadRequest` 对空段落、偏移量、超长文本均有检查）。
- 文件写入采用 **原子写**（`*.tmp` → `Move`）模式，避免配置损坏。
- XML 文档注释完整，公共 API 可读性好。

### 问题与风险

| 优先级 | 问题 | 说明 | 建议 |
|--------|------|------|------|
| **中** | **偏移量计算对多字节字符存在漂移风险** | `FindTextOffset` / `NormalizeText` / `MapNormalizedIndexToOriginal` 基于 UTF-16 `char` 操作。WPS JS API 在 Linux 下的偏移量语义（COM `Range` 基于 `char` 还是字节）未明确，中文文档可能出现“错位一个字”现象。 | 增加单元测试覆盖中文、Emoji、混合空白的段落；必要时引入 `StringInfo` 按文本元素（grapheme）计算。 |
| **中** | **Ollama 超时配置冗余且过长** | `HttpClient.Timeout = 300s`，`SendProofreadMessageAsync` 又创建 300s 的 `CancellationTokenSource`，实际上没有分层超时。Ollama 未启动时，连接建立超时可能高达 5 分钟。 | 分层超时：`HttpClient.Timeout` 设为 310s（兜底），`SendProofreadMessageAsync` 设为 60s~120s（业务预期），`ApiClient.post` 也同步调整。 |
| **低** | **`ProcessTime` 使用 `DateTime.Now` 而非 `DateTime.UtcNow`** | 导致日志和缓存时间带本地时区偏移，不便于跨时区排障。 | 统一改为 `DateTime.UtcNow`。 |
| **低** | **`ConvertToResponse` 中的防御性二次解析** | `result.Items ?? ProofreadIssueParser.ParseProofreadItems(...)` 在控制器中再次执行正则，增加 CPU 开销。 | 既然 `ProcessParagraphAsync` 已解析，可直接信任 `result.Items`，移除 fallback。 |

---

## 4. 健壮性

### 优点
- LLM 调用具备 **指数退避重试**（502/503/504/429/499 + 超时）。
- `CancellationToken` 全程传递，`ProofreadService` 支持 `Dispose` + `RequestCancel`。
- 启动时自动清理旧进程和旧端口文件，避免端口冲突。
- 前端 JS 大量 `try/catch`，WPS API 异常不会导致加载项崩溃。

### 问题与风险

| 优先级 | 问题 | 说明 | 建议 |
|--------|------|------|------|
| **高** | **端口文件清理存在竞态条件** | `Program.cs` 中 `oldProcess.HasExited` 与 `oldProcess.Kill()` 之间可能进程已退出，在部分 Linux 发行版会抛出 `InvalidOperationException` 或 `PlatformNotSupportedException`。 | 将 `Kill()` 和 `WaitForExit()` 整体包入 `try/catch`，或改用 `oldProcess.TryKill()` 扩展方法。 |
| **高** | **`LinuxCrypto.EnsureSalt()` 静默吞掉所有异常** | 若磁盘满或权限错误导致 salt 文件无法写入，函数返回内存中的临时 salt，下次启动 salt 不同，旧的 `config.dat` 将永久无法解密。 | 记录 `Console.Error.WriteLine` 或 `Debug.WriteLine` 日志；若写入失败且旧 salt 不存在，应抛出异常让用户感知。 |
| **中** | **`ApiClient` 对服务发现失败的降级不足** | `wps.FileSystem` 在某些 Linux WPS 版本或沙箱中可能不可用，此时 `discoverService()` 直接失败，没有手动输入端口号的兜底交互。 | UI 增加“后端端口”输入框，当自动发现失败时允许用户手动填写。 |
| **中** | **systemd 检测方式不够可靠** | `systemctl --user status &>/dev/null` 能判断 systemd 是否存在，但无法判断 `user@uid.service` 是否已启动。 | 增加 `loginctl show-user "$USER" --property=State` 或检查 `$XDG_RUNTIME_DIR/systemd` 目录是否存在。 |
| **低** | **`DriveInfo` 在特殊 Linux 文件系统上可能异常** | `ConfigManager.SaveConfig` 中 `new DriveInfo(root.FullName)` 在 overlayfs、tmpfs 或网络挂载点上可能抛出非 `IOException`。 | 当前 `catch (Exception ex) when (ex is not IOException)` 已处理，保持现状即可。 |

---

## 5. 安全性

### 优点
- 随机 256-bit API Token + 固定时间比较，有效防御时序攻击。
- Linux 配置文件采用 **AES-GCM + `/etc/machine-id` + 随机 salt** 加密，优于明文存储。
- 敏感文件设置 `chmod 600`。
- 服务仅绑定 `127.0.0.1`，拒绝远程访问。
- 配置读取接口 **不回显 API Key**。

### 问题与风险

| 优先级 | 问题 | 说明 | 建议 |
|--------|------|------|------|
| **高** | **CORS 策略在 Linux WPS 环境下可能失效** | `WithOrigins("http://localhost:*")` 不是有效的 CORS 通配符，ASP.NET Core 会将其视为字面量，无法匹配 `http://localhost:8080`。若 WPS 在 Linux 使用 WebKit/WebView2 并严格校验 CORS，预检请求会被拒绝。 | 改为 `SetIsOriginAllowed(origin => origin.StartsWith("http://localhost:"))` 或 `AllowAnyOrigin()`（若不需要 Cookie 凭证）。 |
| **中** | **`chmod` 调用存在潜在 Shell 注入** | `ProcessStartInfo("chmod", $"600 \"{path}\"")` 中 `path` 若含双引号会导致参数解析错误（虽然 `ConfigDir` 来自环境变量，通常安全）。 | 使用 `psi.ArgumentList.Add(modeStr)` + `psi.ArgumentList.Add(path)`（.NET 5+ 支持），彻底消除注入风险。 |
| **中** | **`/health` 和 `/config` 免 Token** | 兼容旧版前端的设计，但 `/config` 会泄露使用的模型和提供商信息。 | 首次连接后前端应缓存 Token，后续所有请求均携带；服务端在后续版本中移除 `/config` 的豁免。 |
| **低** | **LLM API Key 在前后端明文传输** | 虽然走本地回环，但其他具备本地 root 权限的进程可抓包 `lo`。 | 在纯本地单用户场景下可接受；若后续支持多用户或远程后端，应引入临时会话密钥加密传输。 |

---

## 6. 性能

### 优点
- `Parallel.ForEachAsync` 替代了原来的 `Task.WhenAny` O(n²) 轮询，并发调度效率显著提升。
- `SharedHttpClientFactory` 复用 `SocketsHttpHandler`，避免 Linux `TIME_WAIT` Socket 耗尽。
- 缓存采用 `ConcurrentDictionary` 降低锁竞争，并设 50MB 内存上限。
- 进度节流（200ms）减少 UI 刷新开销。

### 问题与风险

| 优先级 | 问题 | 说明 | 建议 |
|--------|------|------|------|
| **高** | **缓存内存统计是 O(n×m) 热点** | `StoreResult` → `GetCacheStats()` → `EstimateResultBytes()` 会遍历全部缓存项及其所有子项。若缓存 1000 段、每段 50 处问题，每次存入都要做数万次字符串长度计算。 | 在 `StoreResult` 时直接累加已知的字节数到全局计数器（`Interlocked.Add`），淘汰时减去对应值，将 `GetCacheStats` 降为 O(1)。 |
| **中** | **逐次淘汰 10 条的 while 循环效率低** | 当内存突增远超上限时，`while` 内反复 `EvictOldestEntries(10)` 并重新统计，造成多次字典遍历。 | 改为一次性按需要淘汰的数量（或比例）清理，或改用 `MemoryCache` 内置的 SizeLimit。 |
| **中** | **并发度固定为 5，未区分提供商类型** | Ollama 本地推理通常是 CPU/GPU 密集型，5 并发可能打满本地资源；而 DeepSeek 在线 API 更适合高并发。 | 在 `LLMServiceFactory` 中增加 `RecommendedConcurrency` 属性，Ollama 默认 1~2，在线 API 默认 5。 |
| **低** | **控制器中 `string.Join("\n", request.Paragraphs.Select(...))` 后又分割** | 前端已经分段，后端又重新拼接再分割，造成不必要的字符串分配。 | 在 `ProofreadService` 中增加接受 `List<string>` 段落列表的重载，绕过 `DocumentSegmenter`。 |
| **低** | **未启用响应压缩** | 大量校对结果 JSON 可能达数百 KB，对本地服务影响有限，但启用压缩无额外成本。 | `builder.Services.AddResponseCompression()` + `app.UseResponseCompression()`。 |

---

## 7. Linux 兼容性

### 优点
- 自包含（Self-Contained）+ 单文件发布，无需目标系统预装 .NET 运行时。
- 支持 `linux-x64` 和 `linux-arm64`（含交叉编译探测）。
- 端口文件优先写入 `$XDG_RUNTIME_DIR`，回退到 `~/.config/gowordagent/`，符合 Linux 桌面规范。
- systemd 用户级服务与手动脚本双轨 fallback。
- CSS 字体栈针对中文 Linux 桌面（Noto Sans CJK SC / WenQuanYi Micro Hei）做了优化。
- 脚本对银河麒麟 V10、Ubuntu、RHEL/CentOS 的包管理器均有适配。

### 问题与风险

| 优先级 | 问题 | 说明 | 建议 |
|--------|------|------|------|
| **高** | **WPS 加载项安装路径单一且无 fallback** | 脚本硬编码 `$HOME/.local/share/Kingsoft/wps/jsaddons/...`，但 UOS、Deepin、部分 Kylin 版本可能使用 `/usr/share/wps/office/jsaddons/` 或 `$HOME/.config/Kingsoft/...`。 | 部署脚本增加路径探测逻辑：先检测存在的目录，再复制；UI 给出路径自定义选项。 |
| **中** | **`wps.FileSystem` 在 Linux WPS 中可能不可用** | 这是前端服务发现的核心依赖。部分 Linux WPS 构建未暴露 `FileSystem` API，或沙箱限制文件读取。 | 增加 XMLHttpRequest 轮询 `http://127.0.0.1:默认端口/health` 的 fallback；或在 UI 显示手动端口输入。 |
| **中** | **WPS Linux COM API 对 TrackRevisions / Comments 支持不完整** | 虽然代码有 `try/catch` 兜底，但用户可能发现“应用修订”功能在某些 Linux WPS 版本上完全无效果。 | 在 UI 中增加提示：若应用修订失败，可仅显示结果列表，用户手动修改。 |
| **低** | **卸载脚本使用 `sudo rm -rf /opt/gowordagent`** | `install.sh` 安装到 `/opt`（需要 root），`uninstall.sh` 用 `sudo` 删除是正确的；但如果用户用 `deploy-linux.sh` 安装到 `$HOME/.local/opt`，`uninstall.sh` 不会清理该路径。 | `uninstall.sh` 增加对 `$HOME/.local/opt/gowordagent` 的检查和清理。 |
| **低** | **`chmod` 依赖外部命令** | 在最小化容器或 chroot 环境中可能找不到 `chmod`。 | 可考虑引入 `Mono.Unix.Native.Syscall.chmod` P/Invoke（若不想增加依赖，保持现状也可）。 |

---

## 总体评估

| 维度 | 评分 | 总结 |
|------|------|------|
| 功能完成度 | ⭐⭐⭐⭐☆ | 核心闭环完整，缺少取消、进度流式推送和部分结果保留。 |
| 架构设计 | ⭐⭐⭐⭐☆ | 分层清晰，扩展性好，静态全局状态稍多。 |
| 代码质量 | ⭐⭐⭐⭐☆ | 规范、注释充分，JSON 库混用和偏移量计算有待打磨。 |
| 健壮性 | ⭐⭐⭐⭐☆ | 重试、取消、清理做得不错，竞态和降级场景可再加强。 |
| 安全性 | ⭐⭐⭐⭐☆ | 本地服务场景下足够，CORS 和 chmod 调用有优化空间。 |
| 性能 | ⭐⭐⭐⭐☆ | 并发和连接池优秀，缓存内存统计是明显热点。 |
| Linux 兼容性 | ⭐⭐⭐☆☆ | **最大短板**：WPS 加载项路径和 JS API（`FileSystem`/`TrackRevisions`）在不同 Linux 发行版/WPS 版本上差异大，需要更多 fallback 和探测逻辑。 |

**综合评分**: ⭐⭐⭐⭐☆ (7.5/10)

---

## 建议的优先修复清单

1. **修复 CORS 策略**（高优先级，可能直接导致 Linux WPS 下请求失败）。
2. **优化缓存内存统计**（高优先级，避免大文档时的 O(n×m) 性能塌陷）。
3. **增加后端重启后的 Token 自动刷新**（高优先级，影响可用性）。
4. **增加显式取消端点/机制**（高优先级，减少无效 Token 消耗）。
5. **增加段落级 `try/catch` 以保留部分结果**（中优先级，提升用户体验）。
6. **改进 WPS 加载项路径探测和 `FileSystem` fallback**（中优先级，提升 Linux 兼容性）。
7. **统一 JSON 序列化库**（低优先级，技术债清理）。
8. **修正 `chmod` 的 `ArgumentList` 用法**（低优先级，安全加固）。

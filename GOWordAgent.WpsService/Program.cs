using System.Diagnostics;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading.RateLimiting;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Hosting.Server.Features;
using Microsoft.AspNetCore.RateLimiting;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Serialization;
using GOWordAgentAddIn;
using GOWordAgent.WpsService.Controllers;
using GOWordAgent.WpsService.Auth;
using GOWordAgent.WpsService.Middlewares;

var builder = WebApplication.CreateBuilder(args);

// 配置日志
builder.Logging.ClearProviders();
builder.Logging.AddConsole();
builder.Logging.AddDebug();
builder.Logging.SetMinimumLevel(LogLevel.Information);

// 配置 Kestrel 监听地址
builder.WebHost.ConfigureKestrel(options =>
{
    // 使用 0 让系统自动分配端口
    options.Listen(IPAddress.Loopback, 0);
});

// 注册服务
builder.Services.AddResponseCompression();
builder.Services.AddControllers()
    .AddNewtonsoftJson(options =>
    {
        options.SerializerSettings.ContractResolver = new CamelCasePropertyNamesContractResolver();
        options.SerializerSettings.NullValueHandling = Newtonsoft.Json.NullValueHandling.Ignore;
        options.SerializerSettings.Converters.Add(new StringEnumConverter());
    });

// 请求限流：每 IP 每 60 秒最多 30 次请求（本地服务场景下足够日常使用，防止恶意刷接口）
builder.Services.AddRateLimiter(options =>
{
    options.AddFixedWindowLimiter(policyName: "ProofreadPolicy", opt =>
    {
        opt.PermitLimit = 30;
        opt.Window = TimeSpan.FromMinutes(1);
        opt.QueueProcessingOrder = QueueProcessingOrder.OldestFirst;
        opt.QueueLimit = 2;
    });
    options.OnRejected = (context, token) =>
    {
        context.HttpContext.Response.StatusCode = 429;
        return new ValueTask(context.HttpContext.Response.WriteAsJsonAsync(new { error = "请求过于频繁，请稍后再试" }, token));
    };
});

builder.Services.AddCors(options =>
{
    options.AddPolicy("WpsAddon", policy =>
    {
        policy.SetIsOriginAllowed(origin =>
            {
                // 允许所有 localhost/127.0.0.1 的任意端口（WPS 加载项在不同平台可能使用不同端口）
                return origin.StartsWith("http://localhost:", StringComparison.OrdinalIgnoreCase)
                    || origin.StartsWith("http://127.0.0.1:", StringComparison.OrdinalIgnoreCase)
                    || origin.Equals("http://localhost", StringComparison.OrdinalIgnoreCase)
                    || origin.Equals("http://127.0.0.1", StringComparison.OrdinalIgnoreCase);
            })
            .AllowAnyMethod()
            .AllowAnyHeader()
            .AllowCredentials();
    });
});

var app = builder.Build();

// 加载配置
ConfigManager.LoadConfig();

app.UseResponseCompression();
app.UseCors("WpsAddon");
app.UseRateLimiter();
app.UseMiddleware<RequestLoggingMiddleware>();
app.UseMiddleware<ApiTokenAuthMiddleware>();
app.MapControllers();

// 获取生命周期服务
var lifetime = app.Services.GetRequiredService<IHostApplicationLifetime>();

static string GetPortFilePath()
{
    var runtimeDir = Environment.GetEnvironmentVariable("XDG_RUNTIME_DIR");
    if (!string.IsNullOrEmpty(runtimeDir) && Directory.Exists(runtimeDir))
    {
        return Path.Combine(runtimeDir, $"gowordagent-port-{Environment.UserName}.json");
    }
    var configDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
        ".config", "gowordagent");
    Directory.CreateDirectory(configDir);
    return Path.Combine(configDir, "service-port.json");
}

static void SetUnixFilePermissions(string path, int mode)
{
    if (!RuntimeInformation.IsOSPlatform(OSPlatform.Linux) && !RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        return;
    try
    {
        var psi = new ProcessStartInfo("chmod")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };
        psi.ArgumentList.Add(Convert.ToString(mode, 8));
        psi.ArgumentList.Add(path);
        using var proc = Process.Start(psi);
        proc?.WaitForExit(2000);
    }
    catch { /* 忽略权限设置失败 */ }
}

// 应用启动时清理旧实例并写入端口文件
lifetime.ApplicationStarted.Register(() =>
{
    try
    {
        var portFile = GetPortFilePath();

        // 清理可能存在的旧服务实例
        try
        {
            if (File.Exists(portFile))
            {
                var oldContent = File.ReadAllText(portFile);
                var oldInfo = JsonConvert.DeserializeObject<PortInfo>(oldContent);

                if (oldInfo?.Pid > 0 && oldInfo.Pid != Environment.ProcessId)
                {
                    try
                    {
                        var oldProcess = Process.GetProcessById(oldInfo.Pid);
                        if (oldProcess != null && !oldProcess.HasExited)
                        {
                            try
                            {
                                var processName = oldProcess.ProcessName;
                                if (processName.Contains("gowordagent", StringComparison.OrdinalIgnoreCase))
                                {
                                    Console.WriteLine($"检测到旧服务实例 (PID: {oldInfo.Pid})，正在终止...");
                                    try
                                    {
                                        oldProcess.Kill();
                                        oldProcess.WaitForExit(5000);
                                        Console.WriteLine("旧服务实例已终止");
                                    }
                                    catch (Exception killEx)
                                    {
                                        Console.WriteLine($"终止旧实例时出错: {killEx.Message}");
                                    }
                                }
                            }
                            catch { /* 忽略进程名称获取错误 */ }
                        }
                    }
                    catch (ArgumentException) { /* 进程不存在，忽略 */ }
                }

                File.Delete(portFile);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"清理旧实例时出错: {ex.Message}");
        }

        // 获取服务地址并写入端口文件
        var server = app.Services.GetRequiredService<IServer>();
        var addresses = server.Features.Get<IServerAddressesFeature>()?.Addresses;
        var address = addresses?.FirstOrDefault();

        if (address != null)
        {
            var port = new Uri(address).Port;
            var portInfo = new PortInfo
            {
                Port = port,
                Pid = Environment.ProcessId,
                Timestamp = DateTimeOffset.UtcNow.ToUnixTimeSeconds(),
                ApiToken = ApiTokenAuth.Token
            };

            var tempFile = portFile + ".tmp";
            File.WriteAllText(tempFile, JsonConvert.SerializeObject(portInfo));
            File.Move(tempFile, portFile, overwrite: true);
            SetUnixFilePermissions(portFile, 600);

            Console.WriteLine($"GOWordAgent Service started on port {port}");
            Console.WriteLine($"Port info written to {portFile}");
        }
        else
        {
            Console.Error.WriteLine("Failed to get server address");
        }
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"Failed to write port file: {ex.Message}");
    }
});

// 应用停止时清理端口文件
lifetime.ApplicationStopping.Register(() =>
{
    try
    {
        var portFile = GetPortFilePath();
        if (File.Exists(portFile))
        {
            File.Delete(portFile);
            Console.WriteLine("Port file cleaned up");
        }
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"Failed to cleanup port file: {ex.Message}");
    }
});

// 配置线程池（Linux 环境下根据 CPU 核数动态调整）
int processorCount = Environment.ProcessorCount;
ThreadPool.SetMinThreads(Math.Max(4, processorCount * 2), Math.Max(4, processorCount * 2));
ThreadPool.SetMaxThreads(Math.Min(128, Math.Max(32, processorCount * 4)), Math.Min(128, Math.Max(32, processorCount * 4)));

// 处理 Ctrl+C
Console.CancelKeyPress += (sender, e) =>
{
    e.Cancel = true;
    lifetime.StopApplication();
};

// ASP.NET Core 已内置 SIGTERM 处理，不需要额外注册 ProcessExit
// 保留此注释以提醒后续维护者不要重复添加

app.Run();

/// <summary>
/// 端口信息记录
/// </summary>
public class PortInfo
{
    public int Port { get; set; }
    public int Pid { get; set; }
    public long Timestamp { get; set; }
    public string ApiToken { get; set; } = "";
}

using System.Net;
using GOWordAgentAddIn;
using GOWordAgent.WpsService.Controllers;

var builder = WebApplication.CreateBuilder(args);

// 配置 Kestrel 监听地址
builder.WebHost.ConfigureKestrel(options =>
{
    // 使用 0 让系统自动分配端口
    options.Listen(IPAddress.Loopback, 0);
});

// 注册服务
builder.Services.AddSingleton<ILLMService>(sp =>
{
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

// 加载配置
ConfigManager.LoadConfig();

app.UseCors("WpsAddon");
app.MapControllers();

// 获取实际分配的端口并写入文件
var serverAddresses = app.Services.GetRequiredService<IServer>().Features.Get<IServerAddressesFeature>();
var address = serverAddresses?.Addresses.FirstOrDefault();
if (address != null)
{
    var uri = new Uri(address);
    var port = uri.Port;
    
    var portInfo = new
    {
        port,
        pid = Environment.ProcessId,
        timestamp = DateTimeOffset.UtcNow.ToUnixTimeSeconds()
    };
    
    var portFile = "/tmp/gowordagent-port.json";
    File.WriteAllText(portFile, System.Text.Json.JsonSerializer.Serialize(portInfo));
    Console.WriteLine($"GOWordAgent Service started on port {port}");
    Console.WriteLine($"Port info written to {portFile}");
}

app.Run();

// 清理端口文件
if (File.Exists("/tmp/gowordagent-port.json"))
{
    try { File.Delete("/tmp/gowordagent-port.json"); } catch { }
}

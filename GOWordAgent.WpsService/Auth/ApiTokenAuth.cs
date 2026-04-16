using System;
using System.Security.Cryptography;
using System.Text;
using Microsoft.AspNetCore.Http;
using System.Threading.Tasks;

namespace GOWordAgent.WpsService.Auth
{
    /// <summary>
    /// 本地 API 认证：防止其他本地进程未授权访问
    /// </summary>
    public static class ApiTokenAuth
    {
        private static string? _token;
        private static readonly object _lock = new object();

        public static string Token
        {
            get
            {
                if (_token == null)
                {
                    lock (_lock)
                    {
                        _token ??= GenerateToken();
                    }
                }
                return _token;
            }
        }

        private static string GenerateToken()
        {
            var bytes = RandomNumberGenerator.GetBytes(32);
            return Convert.ToBase64String(bytes);
        }

        /// <summary>
        /// 验证请求是否携带正确的 Token
        /// </summary>
        public static bool Validate(HttpRequest request)
        {
            var header = request.Headers["X-Api-Token"].ToString();
            if (string.IsNullOrEmpty(header))
            {
                // 兼容旧版前端：允许健康检查和配置读取无 Token（首次连接场景）
                var path = request.Path.Value ?? "";
                if (path.EndsWith("/health", StringComparison.OrdinalIgnoreCase) ||
                    path.EndsWith("/config", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
                return false;
            }
            // 固定时间比较，防止时序攻击
            return CryptographicOperations.FixedTimeEquals(
                Encoding.UTF8.GetBytes(header),
                Encoding.UTF8.GetBytes(Token));
        }
    }

    /// <summary>
    /// 认证中间件
    /// </summary>
    public class ApiTokenAuthMiddleware
    {
        private readonly RequestDelegate _next;

        public ApiTokenAuthMiddleware(RequestDelegate next)
        {
            _next = next;
        }

        public async Task InvokeAsync(HttpContext context)
        {
            if (!ApiTokenAuth.Validate(context.Request))
            {
                context.Response.StatusCode = 401;
                await context.Response.WriteAsJsonAsync(new { error = "未授权访问" });
                return;
            }
            await _next(context);
        }
    }
}

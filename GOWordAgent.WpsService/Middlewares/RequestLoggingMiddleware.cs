using System.Diagnostics;

namespace GOWordAgent.WpsService.Middlewares
{
    /// <summary>
    /// 请求日志中间件：记录请求路径、状态码、耗时
    /// </summary>
    public class RequestLoggingMiddleware
    {
        private readonly RequestDelegate _next;
        private readonly ILogger<RequestLoggingMiddleware> _logger;

        public RequestLoggingMiddleware(RequestDelegate next, ILogger<RequestLoggingMiddleware> logger)
        {
            _next = next;
            _logger = logger;
        }

        public async Task InvokeAsync(HttpContext context)
        {
            var stopwatch = Stopwatch.StartNew();
            var path = context.Request.Path;
            var method = context.Request.Method;

            try
            {
                await _next(context);
            }
            finally
            {
                stopwatch.Stop();
                var statusCode = context.Response.StatusCode;
                var level = statusCode >= 500 ? LogLevel.Error : (statusCode >= 400 ? LogLevel.Warning : LogLevel.Information);
                _logger.Log(level, "{Method} {Path} => {StatusCode} in {ElapsedMs}ms", method, path, statusCode, stopwatch.ElapsedMilliseconds);
            }
        }
    }
}

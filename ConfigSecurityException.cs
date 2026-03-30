using System;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// 配置安全异常 - 配置被篡改或损坏时抛出
    /// </summary>
    public class ConfigSecurityException : Exception
    {
        public ConfigSecurityException(string message) : base(message)
        {
        }

        public ConfigSecurityException(string message, Exception innerException) 
            : base(message, innerException)
        {
        }
    }
}

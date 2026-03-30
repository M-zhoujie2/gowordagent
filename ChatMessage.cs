using System;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// 聊天消息类型
    /// </summary>
    public enum ChatRole
    {
        User,      // 用户
        System,    // 系统
        AI         // AI 助手
    }

    /// <summary>
    /// 聊天消息
    /// </summary>
    public class ChatMessage
    {
        /// <summary>
        /// 角色
        /// </summary>
        public ChatRole Role { get; set; }

        /// <summary>
        /// 消息内容
        /// </summary>
        public string Content { get; set; }

        /// <summary>
        /// 时间戳
        /// </summary>
        public DateTime Timestamp { get; set; }

        /// <summary>
        /// 消息是否出错
        /// </summary>
        public bool IsError { get; set; }

        /// <summary>
        /// 创建新消息
        /// </summary>
        public ChatMessage(ChatRole role, string content, bool isError = false)
        {
            Role = role;
            Content = content ?? string.Empty;
            Timestamp = DateTime.Now;
            IsError = isError;
        }

        /// <summary>
        /// 转换为 LLM API 格式的对象
        /// </summary>
        public object ToLLMFormat()
        {
            string roleStr;
            switch (Role)
            {
                case ChatRole.User:
                    roleStr = "user";
                    break;
                case ChatRole.AI:
                    roleStr = "assistant";
                    break;
                case ChatRole.System:
                    roleStr = "system";
                    break;
                default:
                    roleStr = "user";
                    break;
            }
            return new { role = roleStr, content = Content };
        }

        /// <summary>
        /// 显示名称
        /// </summary>
        public string DisplayName
        {
            get
            {
                switch (Role)
                {
                    case ChatRole.User:
                        return "你";
                    case ChatRole.AI:
                        return "AI";
                    case ChatRole.System:
                        return "系统";
                    default:
                        return "未知";
                }
            }
        }
    }
}

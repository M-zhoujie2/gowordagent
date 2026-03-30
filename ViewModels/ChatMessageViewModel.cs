using System;
using System.ComponentModel;
using System.Windows.Input;

namespace GOWordAgentAddIn.ViewModels
{
    /// <summary>
    /// 消息气泡类型
    /// </summary>
    public enum MessageType
    {
        System,
        User,
        AI,
        Error
    }

    /// <summary>
    /// 聊天消息 ViewModel（支持数据绑定）
    /// </summary>
    public class ChatMessageViewModel : INotifyPropertyChanged
    {
        private string _sender;
        private string _content;
        private string _time;
        private MessageType _messageType;
        private bool _showCopyButton;
        private string _copyButtonText = "📋 复制";

        public ChatMessageViewModel()
        {
            _time = DateTime.Now.ToString("HH:mm");
        }

        /// <summary>
        /// 发送者名称
        /// </summary>
        public string Sender
        {
            get => _sender;
            set
            {
                if (_sender != value)
                {
                    _sender = value;
                    OnPropertyChanged(nameof(Sender));
                }
            }
        }

        /// <summary>
        /// 消息内容
        /// </summary>
        public string Content
        {
            get => _content;
            set
            {
                if (_content != value)
                {
                    _content = value;
                    OnPropertyChanged(nameof(Content));
                }
            }
        }

        /// <summary>
        /// 时间显示
        /// </summary>
        public string Time
        {
            get => _time;
            set
            {
                if (_time != value)
                {
                    _time = value;
                    OnPropertyChanged(nameof(Time));
                }
            }
        }

        /// <summary>
        /// 消息类型
        /// </summary>
        public MessageType MessageType
        {
            get => _messageType;
            set
            {
                if (_messageType != value)
                {
                    _messageType = value;
                    OnPropertyChanged(nameof(MessageType));
                    OnPropertyChanged(nameof(IsUser));
                    OnPropertyChanged(nameof(IsError));
                }
            }
        }

        /// <summary>
        /// 是否为用户消息（用于对齐）
        /// </summary>
        public bool IsUser => _messageType == MessageType.User;

        /// <summary>
        /// 是否为错误消息（用于颜色）
        /// </summary>
        public bool IsError => _messageType == MessageType.Error;

        /// <summary>
        /// 是否显示复制按钮
        /// </summary>
        public bool ShowCopyButton
        {
            get => _showCopyButton;
            set
            {
                if (_showCopyButton != value)
                {
                    _showCopyButton = value;
                    OnPropertyChanged(nameof(ShowCopyButton));
                }
            }
        }

        /// <summary>
        /// 复制按钮文本
        /// </summary>
        public string CopyButtonText
        {
            get => _copyButtonText;
            set
            {
                if (_copyButtonText != value)
                {
                    _copyButtonText = value;
                    OnPropertyChanged(nameof(CopyButtonText));
                }
            }
        }

        /// <summary>
        /// 复制命令
        /// </summary>
        public ICommand CopyCommand { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// 创建系统消息
        /// </summary>
        public static ChatMessageViewModel CreateSystem(string message)
        {
            return new ChatMessageViewModel
            {
                Sender = "系统",
                Content = message,
                MessageType = MessageType.System
            };
        }

        /// <summary>
        /// 创建用户消息
        /// </summary>
        public static ChatMessageViewModel CreateUser(string message)
        {
            return new ChatMessageViewModel
            {
                Sender = "你",
                Content = message,
                MessageType = MessageType.User
            };
        }

        /// <summary>
        /// 创建 AI 消息
        /// </summary>
        public static ChatMessageViewModel CreateAI(string message, bool showCopyButton = false)
        {
            return new ChatMessageViewModel
            {
                Sender = "AI",
                Content = message,
                MessageType = MessageType.AI,
                ShowCopyButton = showCopyButton
            };
        }

        /// <summary>
        /// 创建错误消息
        /// </summary>
        public static ChatMessageViewModel CreateError(string message)
        {
            return new ChatMessageViewModel
            {
                Sender = "错误",
                Content = message,
                MessageType = MessageType.Error
            };
        }
    }
}

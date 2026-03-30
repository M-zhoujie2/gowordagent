using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// 聊天视图管理器 - 管理聊天气泡的创建和显示
    /// </summary>
    public class ChatViewManager
    {
        private readonly Panel _messagesPanel;
        private readonly ScrollViewer _scrollViewer;
        private readonly SolidColorBrush _userBubbleColor;
        private readonly SolidColorBrush _aiBubbleColor;
        private readonly SolidColorBrush _textPrimaryColor;
        private readonly SolidColorBrush _textSecondaryColor;

        public ChatViewManager(Panel messagesPanel, ScrollViewer scrollViewer,
            SolidColorBrush userBubbleColor, SolidColorBrush aiBubbleColor,
            SolidColorBrush textPrimaryColor, SolidColorBrush textSecondaryColor)
        {
            _messagesPanel = messagesPanel ?? throw new ArgumentNullException(nameof(messagesPanel));
            _scrollViewer = scrollViewer ?? throw new ArgumentNullException(nameof(scrollViewer));
            _userBubbleColor = userBubbleColor ?? throw new ArgumentNullException(nameof(userBubbleColor));
            _aiBubbleColor = aiBubbleColor ?? throw new ArgumentNullException(nameof(aiBubbleColor));
            _textPrimaryColor = textPrimaryColor ?? throw new ArgumentNullException(nameof(textPrimaryColor));
            _textSecondaryColor = textSecondaryColor ?? throw new ArgumentNullException(nameof(textSecondaryColor));
        }

        /// <summary>
        /// 添加消息气泡到聊天框
        /// </summary>
        public void AddMessageBubble(string sender, string message, bool isUser, bool isError = false)
        {
            try
            {
                if (_messagesPanel == null) return;

                var bubble = MessageBubbleFactory.CreateBubble(
                    sender, message, 
                    isError ? BubbleType.Error : (isUser ? BubbleType.User : BubbleType.AI),
                    copyButton: true);
                
                var container = MessageBubbleFactory.WrapInContainer(bubble, alignRight: isUser);
                
                _messagesPanel.Children.Add(container);
                _scrollViewer.ScrollToEnd();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AddMessageBubble] 错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 添加系统消息
        /// </summary>
        public void AddSystemMessage(string message)
        {
            AddMessageBubble("系统", message, false);
        }

        /// <summary>
        /// 添加用户消息
        /// </summary>
        public void AddUserMessage(string message)
        {
            AddMessageBubble("你", message, true);
        }

        /// <summary>
        /// 添加 AI 消息
        /// </summary>
        public void AddAIMessage(string message, bool copyButton = false)
        {
            AddMessageBubble("AI", message, false);
        }

        /// <summary>
        /// 添加错误消息
        /// </summary>
        public void AddErrorMessage(string message)
        {
            AddMessageBubble("错误", message, false, true);
        }

        /// <summary>
        /// 清空聊天记录
        /// </summary>
        public void ClearMessages()
        {
            _messagesPanel.Children.Clear();
        }
    }
}

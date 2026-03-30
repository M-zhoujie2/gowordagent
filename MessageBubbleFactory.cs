using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// 消息气泡样式类型
    /// </summary>
    public enum BubbleType
    {
        System,
        User,
        AI,
        Error
    }

    /// <summary>
    /// 消息气泡工厂，统一创建各种消息气泡
    /// </summary>
    public static class MessageBubbleFactory
    {
        // 默认颜色配置（冻结的 Brush 提高性能，允许跨线程使用）
        private static readonly SolidColorBrush SystemBubbleColor = CreateFrozenBrush(232, 242, 252);
        private static readonly SolidColorBrush UserBubbleColor = CreateFrozenBrush(227, 242, 253);
        private static readonly SolidColorBrush AIBubbleColor = CreateFrozenBrush(245, 245, 245);
        private static readonly SolidColorBrush ErrorBubbleColor = CreateFrozenBrush(255, 235, 238);
        private static readonly SolidColorBrush ErrorTextColor = CreateFrozenBrush(198, 40, 40);
        private static readonly SolidColorBrush TextPrimaryColor = CreateFrozenBrush(34, 34, 34);
        private static readonly SolidColorBrush TextSecondaryColor = CreateFrozenBrush(153, 153, 153);
        
        /// <summary>
        /// 创建冻结的 SolidColorBrush
        /// </summary>
        private static SolidColorBrush CreateFrozenBrush(byte r, byte g, byte b)
        {
            var brush = new SolidColorBrush(Color.FromRgb(r, g, b));
            brush.Freeze();
            return brush;
        }

        /// <summary>
        /// 创建消息气泡
        /// </summary>
        public static Border CreateBubble(string sender, string message, BubbleType type, bool copyButton = false, Action copyAction = null)
        {
            var style = GetBubbleStyle(type);
            var bgColor = style.Background;
            var cornerRadius = style.CornerRadius;
            
            var bubbleBorder = new Border
            {
                Background = bgColor,
                CornerRadius = cornerRadius,
                Padding = new Thickness(12, 10, 12, 10),
                MaxWidth = 350,
                Margin = new Thickness(0, 0, 0, 4)
            };

            var mainStack = new StackPanel();

            // 头部：发送者 + 时间 (+ 复制按钮)
            var headerPanel = new Grid();
            headerPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            
            var headerText = new TextBlock
            {
                Text = string.Format("{0} {1:HH:mm}", sender, DateTime.Now),
                FontSize = 10,
                Foreground = TextSecondaryColor,
                VerticalAlignment = VerticalAlignment.Center
            };
            headerPanel.Children.Add(headerText);

            // 复制按钮（可选）
            if (copyButton && copyAction != null)
            {
                headerPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                
                var copyBtn = new Button
                {
                    Content = "📋 复制",
                    FontSize = 9,
                    Foreground = TextSecondaryColor,
                    Background = Brushes.Transparent,
                    BorderThickness = new Thickness(0),
                    Padding = new Thickness(4, 0, 0, 0),
                    Cursor = Cursors.Hand,
                    VerticalAlignment = VerticalAlignment.Center
                };
                
                copyBtn.Click += (s, e) =>
                {
                    try
                    {
                        copyAction();
                        copyBtn.Content = "✓ 已复制";
                        
                        // 2秒后恢复
                        var timer = new System.Windows.Threading.DispatcherTimer 
                        { 
                            Interval = TimeSpan.FromSeconds(2) 
                        };
                        EventHandler tickHandler = null;
                        tickHandler = (ts, te) =>
                        {
                            copyBtn.Content = "📋 复制";
                            timer.Stop();
                            timer.Tick -= tickHandler;
                        };
                        timer.Tick += tickHandler;
                        timer.Start();
                    }
                    catch { }
                };
                
                Grid.SetColumn(copyBtn, 1);
                headerPanel.Children.Add(copyBtn);
            }

            mainStack.Children.Add(headerPanel);

            // 消息内容
            var contentText = new TextBlock
            {
                Text = message,
                FontSize = 13,
                Foreground = type == BubbleType.Error ? ErrorTextColor : TextPrimaryColor,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 6, 0, 0)
            };
            mainStack.Children.Add(contentText);

            bubbleBorder.Child = mainStack;
            return bubbleBorder;
        }

        /// <summary>
        /// 创建系统消息气泡
        /// </summary>
        public static Border CreateSystemBubble(string message)
        {
            return CreateBubble("系统", message, BubbleType.System);
        }

        /// <summary>
        /// 创建用户消息气泡
        /// </summary>
        public static Border CreateUserBubble(string message)
        {
            return CreateBubble("你", message, BubbleType.User);
        }

        /// <summary>
        /// 创建 AI 消息气泡
        /// </summary>
        public static Border CreateAIBubble(string message, bool copyButton = false)
        {
            if (copyButton)
                return CreateBubble("AI", message, BubbleType.AI, true, () => Clipboard.SetText(message));
            else
                return CreateBubble("AI", message, BubbleType.AI);
        }

        /// <summary>
        /// 创建错误消息气泡
        /// </summary>
        public static Border CreateErrorBubble(string message)
        {
            return CreateBubble("错误", message, BubbleType.Error);
        }

        /// <summary>
        /// 将气泡包装为容器
        /// </summary>
        public static Grid WrapInContainer(Border bubble, bool alignRight = false)
        {
            return new Grid
            {
                HorizontalAlignment = alignRight ? HorizontalAlignment.Right : HorizontalAlignment.Left,
                Margin = new Thickness(0, 0, 0, 8),
                Children = { bubble }
            };
        }

        /// <summary>
        /// 获取气泡样式
        /// </summary>
        private static BubbleStyle GetBubbleStyle(BubbleType type)
        {
            switch (type)
            {
                case BubbleType.System:
                    return new BubbleStyle { Background = SystemBubbleColor, CornerRadius = new CornerRadius(8) };
                case BubbleType.User:
                    return new BubbleStyle { Background = UserBubbleColor, CornerRadius = new CornerRadius(12, 12, 4, 12) };
                case BubbleType.Error:
                    return new BubbleStyle { Background = ErrorBubbleColor, CornerRadius = new CornerRadius(8) };
                case BubbleType.AI:
                default:
                    return new BubbleStyle { Background = AIBubbleColor, CornerRadius = new CornerRadius(12, 12, 12, 4) };
            }
        }

        private struct BubbleStyle
        {
            public SolidColorBrush Background { get; set; }
            public CornerRadius CornerRadius { get; set; }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using GOWordAgentAddIn.Models;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// 校对结果渲染器 - 负责校对结果的 UI 渲染
    /// </summary>
    public class ProofreadResultRenderer
    {
        // 静态复用的画刷 - 避免高频创建导致 GC 压力
        private static readonly SolidColorBrush _separatorBrush = CreateFrozenBrush(200, 200, 200);
        private static readonly SolidColorBrush _buttonBackgroundBrush = CreateFrozenBrush(250, 250, 250);
        private static readonly SolidColorBrush _buttonBorderBrush = CreateFrozenBrush(220, 220, 220);
        private static readonly SolidColorBrush _innerSeparatorBrush = CreateFrozenBrush(230, 230, 230);
        private static readonly SolidColorBrush _modifiedTextBrush = CreateFrozenBrush(0, 120, 0);
        private static readonly SolidColorBrush _modifiedLabelBrush = CreateFrozenBrush(0, 150, 0);
        private static readonly SolidColorBrush _severityMediumBrush = CreateFrozenBrush(255, 152, 0);

        private readonly Panel _messagesPanel;
        private readonly ScrollViewer _scrollViewer;
        private readonly SolidColorBrush _aiBubbleColor;
        private readonly SolidColorBrush _textPrimaryColor;
        private readonly SolidColorBrush _textSecondaryColor;
        private readonly Action<ProofreadIssueItem> _navigateAction;

        /// <summary>
        /// 创建冻结的画刷（线程安全，可复用）
        /// </summary>
        private static SolidColorBrush CreateFrozenBrush(byte r, byte g, byte b)
        {
            var brush = new SolidColorBrush(Color.FromRgb(r, g, b));
            brush.Freeze();
            return brush;
        }

        public ProofreadResultRenderer(Panel messagesPanel, ScrollViewer scrollViewer,
            SolidColorBrush aiBubbleColor, SolidColorBrush textPrimaryColor, SolidColorBrush textSecondaryColor,
            Action<ProofreadIssueItem> navigateAction)
        {
            _messagesPanel = messagesPanel ?? throw new ArgumentNullException(nameof(messagesPanel));
            _scrollViewer = scrollViewer ?? throw new ArgumentNullException(nameof(scrollViewer));
            _aiBubbleColor = aiBubbleColor ?? throw new ArgumentNullException(nameof(aiBubbleColor));
            _textPrimaryColor = textPrimaryColor ?? throw new ArgumentNullException(nameof(textPrimaryColor));
            _textSecondaryColor = textSecondaryColor ?? throw new ArgumentNullException(nameof(textSecondaryColor));
            _navigateAction = navigateAction ?? throw new ArgumentNullException(nameof(navigateAction));
        }

        /// <summary>
        /// 添加校对结果气泡
        /// </summary>
        public void AddProofreadResultBubble(string reportTitle, string reportContent, 
            List<ProofreadIssueItem> items, List<ParagraphResult> paragraphResults)
        {
            var bubbleBorder = new Border
            {
                Background = _aiBubbleColor,
                CornerRadius = new CornerRadius(12, 12, 12, 4),
                Padding = new Thickness(12, 10, 12, 10),
                MaxWidth = 350,
                Margin = new Thickness(0, 0, 0, 4)
            };

            var mainStack = new StackPanel();
            
            // 头部：发送者 + 时间 + 复制按钮
            var headerPanel = CreateHeaderPanel(reportTitle, reportContent);
            mainStack.Children.Add(headerPanel);

            // 报告内容
            var reportText = new TextBlock
            {
                Text = reportContent,
                FontSize = 13,
                Foreground = _textPrimaryColor,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 6, 0, 8)
            };
            mainStack.Children.Add(reportText);

            // 添加问题列表
            if (items.Count > 0)
            {
                AddIssueList(mainStack, items);
            }

            bubbleBorder.Child = mainStack;

            var container = new Grid
            {
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(0, 0, 0, 8)
            };
            container.Children.Add(bubbleBorder);

            _messagesPanel.Children.Add(container);
            _scrollViewer.ScrollToEnd();
        }

        /// <summary>
        /// 创建头部面板
        /// </summary>
        private Grid CreateHeaderPanel(string reportTitle, string reportContent)
        {
            var headerPanel = new Grid();
            headerPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            
            var headerText = new TextBlock
            {
                Text = $"{reportTitle} {DateTime.Now:HH:mm}",
                FontSize = 10,
                Foreground = _textSecondaryColor,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(headerText, 0);
            headerPanel.Children.Add(headerText);
            
            // 复制按钮
            var copyButton = new Button
            {
                Content = "📋 复制",
                FontSize = 9,
                Foreground = _textSecondaryColor,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Padding = new Thickness(4, 0, 0, 0),
                Cursor = Cursors.Hand,
                VerticalAlignment = VerticalAlignment.Center
            };
            copyButton.Click += (s, e) =>
            {
                try
                {
                    Clipboard.SetText(reportContent);
                    copyButton.Content = "✓ 已复制";
                    
                    var timer = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
                    EventHandler tickHandler = null;
                    tickHandler = (ts, te) =>
                    {
                        copyButton.Content = "📋 复制";
                        timer.Stop();
                        timer.Tick -= tickHandler;
                    };
                    timer.Tick += tickHandler;
                    timer.Start();
                }
                catch { }
            };
            Grid.SetColumn(copyButton, 1);
            headerPanel.Children.Add(copyButton);
            
            return headerPanel;
        }

        /// <summary>
        /// 添加问题列表
        /// </summary>
        private void AddIssueList(StackPanel mainStack, List<ProofreadIssueItem> items)
        {
            // 添加分割线
            mainStack.Children.Add(new Separator 
            { 
                Margin = new Thickness(0, 8, 0, 4),
                Background = _separatorBrush
            });
            
            // 问题列表标题
            var listTitle = new TextBlock
            {
                Text = $"📍 共 {items.Count} 处问题（点击定位到文档）：",
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = _textPrimaryColor,
                Margin = new Thickness(0, 4, 0, 6)
            };
            mainStack.Children.Add(listTitle);

            // 添加可点击的问题项
            int displayCount = Math.Min(items.Count, 10);
            for (int i = 0; i < displayCount; i++)
            {
                var item = items[i];
                var itemButton = CreateIssueButton(item, i + 1);
                mainStack.Children.Add(itemButton);
            }
            
            if (items.Count > 10)
            {
                var moreText = new TextBlock
                {
                    Text = $"... 还有 {items.Count - 10} 处问题",
                    FontSize = 11,
                    Foreground = _textSecondaryColor,
                    Margin = new Thickness(4, 4, 0, 0)
                };
                mainStack.Children.Add(moreText);
            }
        }

        /// <summary>
        /// 创建问题项按钮
        /// </summary>
        private Button CreateIssueButton(ProofreadIssueItem item, int displayIndex)
        {
            var button = new Button
            {
                Background = _buttonBackgroundBrush,
                BorderBrush = _buttonBorderBrush,
                BorderThickness = new Thickness(1),
                Padding = new Thickness(10, 8, 10, 8),
                Margin = new Thickness(0, 0, 0, 6),
                HorizontalAlignment = HorizontalAlignment.Stretch,
                HorizontalContentAlignment = HorizontalAlignment.Left,
                Cursor = Cursors.Hand
            };

            var contentStack = new StackPanel();
            
            // 问题类型和序号行
            var headerPanel = new StackPanel { Orientation = Orientation.Horizontal };
            
            // 根据严重程度设置颜色
            Brush severityBrush = GetSeverityBrush(item.Severity);
            
            var indexText = new TextBlock
            {
                Text = $"[{displayIndex}] ",
                FontSize = 11,
                FontWeight = FontWeights.Bold,
                Foreground = severityBrush
            };
            headerPanel.Children.Add(indexText);
            
            var typeText = new TextBlock
            {
                Text = item.Type,
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = _textPrimaryColor
            };
            headerPanel.Children.Add(typeText);
            
            if (!string.IsNullOrEmpty(item.Severity))
            {
                var severityText = new TextBlock
                {
                    Text = $" ({item.Severity})",
                    FontSize = 10,
                    Foreground = severityBrush
                };
                headerPanel.Children.Add(severityText);
            }
            
            contentStack.Children.Add(headerPanel);
            
            // 添加分割线
            contentStack.Children.Add(new Separator 
            { 
                Margin = new Thickness(0, 4, 0, 4),
                Background = _innerSeparatorBrush
            });
            
            // 原文
            AddLabeledText(contentStack, "原文：", item.Original, _textSecondaryColor, 8);
            
            // 修改
            AddLabeledText(contentStack, "修改：", item.Modified, _modifiedTextBrush, 8, true);
            
            // 理由
            AddLabeledText(contentStack, "理由：", item.Reason, _textSecondaryColor, 0);
            
            button.Content = contentStack;
            
            // 点击事件 - 定位到文档
            button.Click += (s, e) => _navigateAction(item);
            
            return button;
        }

        /// <summary>
        /// 获取严重程度对应的颜色
        /// </summary>
        private Brush GetSeverityBrush(string severity)
        {
            if (string.IsNullOrEmpty(severity))
                return Brushes.Gray;
                
            if (severity.Contains("高") || severity.Contains("严重"))
                return Brushes.Red;
            else if (severity.Contains("中"))
                return _severityMediumBrush;
            else if (severity.Contains("低"))
                return Brushes.Green;
                
            return Brushes.Gray;
        }

        /// <summary>
        /// 添加带标签的文本
        /// </summary>
        private void AddLabeledText(StackPanel parent, string label, string text, 
            Brush labelColor, int bottomMargin, bool isModified = false)
        {
            var labelBlock = new TextBlock
            {
                Text = label,
                FontSize = 10,
                FontWeight = FontWeights.SemiBold,
                Foreground = isModified ? _modifiedLabelBrush : labelColor
            };
            parent.Children.Add(labelBlock);
            
            var textBlock = new TextBlock
            {
                Text = text,
                FontSize = isModified ? 11 : (label == "理由：" ? 10 : 11),
                Foreground = isModified ? _modifiedTextBrush : 
                    (label == "理由：" ? labelColor : _textPrimaryColor),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(8, 0, 0, bottomMargin)
            };
            parent.Children.Add(textBlock);
        }
    }
}

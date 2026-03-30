using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;

namespace GOWordAgentAddIn.ViewModels
{
    /// <summary>
    /// 布尔值转换为水平对齐方式（左/右）
    /// </summary>
    public class BoolToAlignmentConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (value is bool isUser && isUser) ? HorizontalAlignment.Right : HorizontalAlignment.Left;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// 布尔值转换为圆角半径（用户消息：右下圆角，AI消息：左下圆角）
    /// </summary>
    public class BoolToCornerRadiusConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isUser && isUser)
                return new CornerRadius(12, 12, 4, 12);  // 用户：右下尖角
            else
                return new CornerRadius(12, 12, 12, 4);  // AI：左下尖角
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// 消息类型转换为背景画刷
    /// </summary>
    public class MessageTypeToBrushConverter : IValueConverter
    {
        private static readonly SolidColorBrush SystemBrush = CreateFrozenBrush(232, 242, 252);
        private static readonly SolidColorBrush UserBrush = CreateFrozenBrush(227, 242, 253);
        private static readonly SolidColorBrush AIBrush = CreateFrozenBrush(245, 245, 245);
        private static readonly SolidColorBrush ErrorBrush = CreateFrozenBrush(255, 235, 238);

        private static SolidColorBrush CreateFrozenBrush(byte r, byte g, byte b)
        {
            var brush = new SolidColorBrush(Color.FromRgb(r, g, b));
            brush.Freeze();
            return brush;
        }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is MessageType type)
            {
                switch (type)
                {
                    case MessageType.System:
                        return SystemBrush;
                    case MessageType.User:
                        return UserBrush;
                    case MessageType.Error:
                        return ErrorBrush;
                    case MessageType.AI:
                    default:
                        return AIBrush;
                }
            }
            return AIBrush;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// 布尔值转换为错误文本画刷
    /// </summary>
    public class BoolToErrorBrushConverter : IValueConverter
    {
        private static readonly SolidColorBrush NormalBrush = CreateFrozenBrush(34, 34, 34);
        private static readonly SolidColorBrush ErrorBrush = CreateFrozenBrush(198, 40, 40);

        private static SolidColorBrush CreateFrozenBrush(byte r, byte g, byte b)
        {
            var brush = new SolidColorBrush(Color.FromRgb(r, g, b));
            brush.Freeze();
            return brush;
        }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (value is bool isError && isError) ? ErrorBrush : NormalBrush;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}

using System;
using System.ComponentModel;
using System.Windows;

namespace GOWordAgentAddIn.ViewModels
{
    /// <summary>
    /// 复杂消息 ViewModel（支持自定义 UI 内容）
    /// </summary>
    public class ComplexMessageViewModel : INotifyPropertyChanged
    {
        private FrameworkElement _content;
        private bool _alignRight;

        /// <summary>
        /// 自定义 UI 内容
        /// </summary>
        public FrameworkElement Content
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
        /// 是否右对齐
        /// </summary>
        public bool AlignRight
        {
            get => _alignRight;
            set
            {
                if (_alignRight != value)
                {
                    _alignRight = value;
                    OnPropertyChanged(nameof(AlignRight));
                    OnPropertyChanged(nameof(HorizontalAlignment));
                }
            }
        }

        /// <summary>
        /// 水平对齐方式
        /// </summary>
        public HorizontalAlignment HorizontalAlignment => 
            _alignRight ? HorizontalAlignment.Right : HorizontalAlignment.Left;

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

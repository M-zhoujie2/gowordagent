using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using GOWordAgentAddIn.Models;
using GOWordAgentAddIn.ViewModels;
using Word = Microsoft.Office.Interop.Word;

namespace GOWordAgentAddIn
{
    public partial class GOWordAgentPaneWpf : UserControl
    {
        private ILLMService _llmService;
        private List<ChatMessage> _messageHistory = new List<ChatMessage>();
        private AIProvider _currentProvider = AIProvider.DeepSeek;
        private readonly WordProofreadController _wordController = new WordProofreadController();
        
        /// <summary>
        /// 聊天消息集合（用于数据绑定，支持多种 ViewModel 类型）
        /// </summary>
        public ObservableCollection<object> ChatMessages { get; } = new ObservableCollection<object>();

        private static readonly SolidColorBrush _primaryColor = CreateFrozenBrush(0, 120, 212);
        private static readonly SolidColorBrush _primaryLightColor = CreateFrozenBrush(232, 242, 252);
        private static readonly SolidColorBrush _userBubbleColor = CreateFrozenBrush(227, 242, 253);
        private static readonly SolidColorBrush _aiBubbleColor = CreateFrozenBrush(245, 245, 245);
        private static readonly SolidColorBrush _textPrimaryColor = CreateFrozenBrush(34, 34, 34);
        private static readonly SolidColorBrush _textSecondaryColor = CreateFrozenBrush(153, 153, 153);
        
        
        /// <summary>
        /// 创建冻结的 SolidColorBrush（提高性能，允许跨线程使用）
        /// </summary>
        private static SolidColorBrush CreateFrozenBrush(byte r, byte g, byte b)
        {
            var brush = new SolidColorBrush(Color.FromRgb(r, g, b));
            brush.Freeze();
            return brush;
        }

        public GOWordAgentPaneWpf()
        {
            InitializeComponent();
            
            // 设置数据绑定
            ChatMessagesControl.ItemsSource = ChatMessages;
            
            InitializeEvents();
            InitializeAIProviderComboBox();
            LoadSavedConfig();
            // 注意：CmbProofreadMode.SelectionChanged 事件在 LoadSavedConfig 中已经注册
            
            // 设置顶部标题栏为就绪状态
            UpdateHeaderStatus("就绪", Brushes.Green);
            
            // 显示初始化消息（直接显示，不延迟）
            string proofreadMode = ConfigManager.CurrentConfig.ProofreadMode ?? "精准校验";
            AddMessageBubble("系统", $"当前校验模式 {proofreadMode}", false);
            
            // 检查是否自动连接
            if (ConfigManager.CurrentConfig.AutoConnect)
            {
                _ = AutoConnectAsyncWithMessage();
            }
            else
            {
                UpdateAIConfigStatus("未连接", _textSecondaryColor);
                AddMessageBubble("系统", "当前未连接AI大模型", false);
            }
        }

        /// <summary>
        /// 自动连接并显示消息
        /// </summary>
        private async Task AutoConnectAsyncWithMessage()
        {
            try
            {
                var config = ConfigManager.CurrentConfig;
                var providerConfig = ConfigManager.GetProviderConfig(config.Provider);
                
                if (providerConfig == null || string.IsNullOrWhiteSpace(providerConfig.ApiKey))
                {
                    UpdateStatus("未找到保存的 API Key，请手动连接", _textSecondaryColor);
                    AddMessageBubble("系统", "当前未连接AI大模型", false);
                    return;
                }

                UpdateStatus("正在自动连接...", _primaryColor);
                
                _llmService = LLMServiceFactory.CreateService(
                    config.Provider, 
                    providerConfig.ApiKey, 
                    providerConfig.ApiUrl, 
                    providerConfig.Model);

                UpdateStatus($"已连接到 {_llmService.ProviderName}", Brushes.Green);
                
                // 显示AI模型连接信息（直接显示）
                AddMessageBubble("系统", $"当前连接的AI大模型 {_llmService.ProviderName}", false);
            }
            catch (Exception ex)
            {
                UpdateStatus($"自动连接失败: {ex.Message}", Brushes.Red);
                AddMessageBubble("系统", "当前未连接AI大模型", false);
            }
        }

        private void InitializeEvents()
        {
            BtnNavChat.Click += (s, e) => ShowChatView();
            BtnNavSettings.Click += (s, e) => ShowSettingsView();
            BtnSend.Click += BtnSend_Click;
            BtnConnect.Click += BtnConnect_Click;
            BtnProofread.Click += BtnProofread_Click;
            BtnClear.Click += BtnClear_Click;
            BtnSaveProofreadConfig.Click += BtnSaveProofreadConfig_Click;
            CmbAIProvider.SelectionChanged += CmbAIProvider_SelectionChanged;
            
            // 初始化校验配置
            InitializeProofreadConfig();
            
            TxtInput.GotFocus += (s, e) => {
                if (TxtInput.Text == "输入消息...")
                {
                    TxtInput.Text = "";
                    TxtInput.Foreground = _textPrimaryColor;
                }
            };
            TxtInput.LostFocus += (s, e) => {
                if (string.IsNullOrWhiteSpace(TxtInput.Text))
                {
                    TxtInput.Text = "输入消息...";
                    TxtInput.Foreground = _textSecondaryColor;
                }
            };
        }

        private void ShowChatView()
        {
            ChatScrollViewer.Visibility = Visibility.Visible;
            SettingsScrollViewer.Visibility = Visibility.Collapsed;
            BtnNavChat.Background = _primaryLightColor;
            SetNavButtonStyle(BtnNavChat, true);
            SetNavButtonStyle(BtnNavSettings, false);
        }

        private void ShowSettingsView()
        {
            ChatScrollViewer.Visibility = Visibility.Collapsed;
            SettingsScrollViewer.Visibility = Visibility.Visible;
            BtnNavSettings.Background = _primaryLightColor;
            SetNavButtonStyle(BtnNavSettings, true);
            SetNavButtonStyle(BtnNavChat, false);
        }

        private void SetNavButtonStyle(Button btn, bool isActive)
        {
            var panel = btn.Content as StackPanel;
            if (panel == null) return;
            var path = panel.Children[0] as Path;
            var text = panel.Children[1] as TextBlock;
            var brush = isActive ? _primaryColor : _textSecondaryColor;
            if (path != null) path.Fill = brush;
            if (text != null) text.Foreground = brush;
        }

        private void InitializeAIProviderComboBox()
        {
            CmbAIProvider.Items.Clear();
            foreach (var provider in LLMServiceFactory.GetProviders())
                CmbAIProvider.Items.Add(new ProviderItem { Provider = provider.Key, Name = provider.Value });
            
            CmbAIProvider.DisplayMemberPath = "Name";
            if (CmbAIProvider.Items.Count > 0) CmbAIProvider.SelectedIndex = 0;
        }

        private void LoadSavedConfig()
        {
            try
            {
                ConfigManager.LoadConfig();
                var config = ConfigManager.CurrentConfig;

                foreach (ProviderItem item in CmbAIProvider.Items)
                {
                    if (item.Provider == config.Provider)
                    {
                        CmbAIProvider.SelectedItem = item;
                        break;
                    }
                }

                var lastCfg = ConfigManager.GetProviderConfig(config.Provider);
                if (lastCfg != null)
                {
                    TxtApiKey.Password = lastCfg.ApiKey ?? "";
                    TxtApiUrl.Text = lastCfg.ApiUrl ?? "";
                    TxtModel.Text = lastCfg.Model ?? "";
                }
                else
                {
                    TxtApiKey.Password = config.ApiKey ?? "";
                    TxtApiUrl.Text = config.ApiUrl ?? "";
                    TxtModel.Text = config.Model ?? "";
                }

                UpdateUIForProvider(config.Provider, onlyFillEmpty: true);
                
                // 加载校验配置（先解绑事件，避免触发 SelectionChanged 覆盖提示词）
                CmbProofreadMode.SelectionChanged -= CmbProofreadMode_SelectionChanged;
                try
                {
                    if (CmbProofreadMode != null)
                    {
                        string savedMode = ConfigManager.CurrentConfig.ProofreadMode ?? "精准校验";
                        CmbProofreadMode.SelectedIndex = savedMode == "全文校验" ? 1 : 0;
                    }
                    if (TxtProofreadPrompt != null)
                    {
                        // 根据当前模式加载对应的提示词
                        string mode = ConfigManager.CurrentConfig.ProofreadMode ?? "精准校验";
                        string savedPrompt = ConfigManager.GetProofreadPromptForMode(mode);
                        if (!string.IsNullOrEmpty(savedPrompt))
                        {
                            TxtProofreadPrompt.Text = savedPrompt;
                        }
                        else
                        {
                            // 如果没有保存过提示词，根据模式设置默认提示词
                            TxtProofreadPrompt.Text = mode == "全文校验" ? GetFullTextProofreadPrompt() : GetPreciseProofreadPrompt();
                        }
                    }
                    // 更新状态显示
                    string currentMode = ConfigManager.CurrentConfig.ProofreadMode ?? "精准校验";
                    if (LblProofreadStatus != null)
                    {
                        LblProofreadStatus.Text = $"状态: 当前为{currentMode}";
                        LblProofreadStatus.Foreground = Brushes.Green;
                    }
                }
                finally
                {
                    CmbProofreadMode.SelectionChanged += CmbProofreadMode_SelectionChanged;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"加载配置失败: {ex.Message}");
            }
        }

        private void CmbAIProvider_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CmbAIProvider.SelectedItem is ProviderItem item)
            {
                _currentProvider = item.Provider;
                
                // 切换下拉框只更新界面配置，不断开当前连接
                // AI 问答继续使用已连接的模型，直到点击"保存并连接"
                var saved = ConfigManager.GetProviderConfig(_currentProvider);
                if (saved != null)
                {
                    TxtApiKey.Password = saved.ApiKey ?? "";
                    TxtApiUrl.Text = saved.ApiUrl ?? "";
                    TxtModel.Text = saved.Model ?? "";
                }
                else
                {
                    // 只更新UI填充默认值，不断开当前连接
                    if (TxtApiUrl != null)
                        TxtApiUrl.Text = LLMServiceFactory.GetDefaultApiUrl(_currentProvider);
                    if (TxtModel != null)
                        TxtModel.Text = LLMServiceFactory.GetDefaultModel(_currentProvider);
                    if (TxtApiKey != null)
                        TxtApiKey.Password = string.Empty;
                }
                
                // 显示当前连接状态
                if (_llmService != null)
                {
                    UpdateStatus($"已选择 {item.Name}（当前连接: {_llmService.ProviderName}），点击保存并连接可切换", _textSecondaryColor);
                }
                else
                {
                    UpdateStatus($"已选择 {item.Name}，请点击保存并连接", _textSecondaryColor);
                }
            }
        }

        private void UpdateUIForProvider(AIProvider provider, bool onlyFillEmpty = true, bool clearApiKey = false)
        {
            if (TxtApiUrl != null && (!onlyFillEmpty || string.IsNullOrWhiteSpace(TxtApiUrl.Text)))
                TxtApiUrl.Text = LLMServiceFactory.GetDefaultApiUrl(provider);

            if (TxtModel != null && (!onlyFillEmpty || string.IsNullOrWhiteSpace(TxtModel.Text)))
                TxtModel.Text = LLMServiceFactory.GetDefaultModel(provider);

            if (clearApiKey && TxtApiKey != null)
                TxtApiKey.Password = string.Empty;

            _llmService = null;
            UpdateStatus($"已选择 {LLMServiceFactory.GetProviders()[provider]}，请点击连接", _textSecondaryColor);
        }

        private async void BtnConnect_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await BtnConnectClickInternal();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[BtnConnect_Click] 未捕获异常: {ex}");
                UpdateStatus("连接失败", Brushes.Red);
            }
        }
        
        private async Task BtnConnectClickInternal()
        {
            string apiKey = TxtApiKey.Password?.Trim();
            if (string.IsNullOrWhiteSpace(apiKey))
            {
                MessageBox.Show("请输入 API Key", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            UpdateStatus("正在连接...", _primaryColor);
            
            try
            {
                _llmService = LLMServiceFactory.CreateService(_currentProvider, apiKey, TxtApiUrl.Text.Trim(), TxtModel.Text.Trim());
                string providerName = LLMServiceFactory.GetProviders()[_currentProvider];
                UpdateStatus($"已连接到 {providerName}，AI 问答将使用此模型", Brushes.Green);
                AddMessageBubble("系统", $"已连接到 {providerName}，现在可以在 AI 问答中使用", false);
                ConfigManager.SaveCurrentConfig(_currentProvider, apiKey, TxtApiUrl.Text.Trim(), TxtModel.Text.Trim(), autoConnect: true);
            }
            catch (Exception ex)
            {
                UpdateStatus($"连接失败 - {ex.Message}", Brushes.Red);
            }
        }

        private async void BtnSend_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await BtnSendClickInternal();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[BtnSend_Click] 未捕获异常: {ex}");
                AddMessageBubble("错误", "发送消息失败，请重试", false, true);
                BtnSend.IsEnabled = true;
            }
        }
        
        private async Task BtnSendClickInternal()
        {
            if (_llmService == null)
            {
                MessageBox.Show("请先点击「保存并连接」按钮连接大模型", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string userMessage = TxtInput.Text.Trim();
            if (string.IsNullOrWhiteSpace(userMessage) || userMessage == "输入消息...")
            {
                MessageBox.Show("请输入消息", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            AddMessageBubble("用户", userMessage, true);
            TxtInput.Text = "输入消息...";
            TxtInput.Foreground = _textSecondaryColor;
            BtnSend.IsEnabled = false;
            UpdateStatus($"{_llmService.ProviderName} 正在思考...", _primaryColor);

            _messageHistory.Add(new ChatMessage(ChatRole.User, userMessage));

            try
            {
                // 转换为 LLM API 格式
                var messagesForLLM = _messageHistory.ConvertAll(m => m.ToLLMFormat());
                string response = await _llmService.SendMessagesWithHistoryAsync(messagesForLLM);
                AddMessageBubble(_llmService.ProviderName, response, false);
                _messageHistory.Add(new ChatMessage(ChatRole.AI, response));
                UpdateStatus("就绪", Brushes.Green);
            }
            catch (Exception ex)
            {
                AddMessageBubble("错误", ex.Message, false, true);
            }
            finally
            {
                BtnSend.IsEnabled = true;
            }
        }

        private void AddMessageBubble(string sender, string message, bool isUser, bool isError = false)
        {
            try
            {
                ChatMessageViewModel vm;
                if (isError)
                    vm = ChatMessageViewModel.CreateError(message);
                else if (isUser)
                    vm = ChatMessageViewModel.CreateUser(message);
                else if (sender == "系统")
                    vm = ChatMessageViewModel.CreateSystem(message);
                else
                    vm = ChatMessageViewModel.CreateAI(message, showCopyButton: true);
                
                // 设置复制命令
                if (vm.ShowCopyButton)
                {
                    vm.CopyCommand = new RelayCommand(() =>
                    {
                        try
                        {
                            Clipboard.SetText(vm.Content);
                            vm.CopyButtonText = "✓ 已复制";
                            
                            var timer = new System.Windows.Threading.DispatcherTimer 
                            { 
                                Interval = TimeSpan.FromSeconds(2) 
                            };
                            EventHandler tickHandler = null;
                            tickHandler = (ts, te) =>
                            {
                                vm.CopyButtonText = "📋 复制";
                                timer.Stop();
                                timer.Tick -= tickHandler;
                            };
                            timer.Tick += tickHandler;
                            timer.Start();
                        }
                        catch { }
                    });
                }
                
                ChatMessages.Add(vm);
                ChatScrollViewer.ScrollToEnd();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AddMessageBubble] 错误: {ex.Message}");
            }
        }

        private void BtnProofread_Click(object sender, RoutedEventArgs e)
        {
            StartProofread();
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            // 清空聊天框
            ChatMessages.Clear();
            _messageHistory.Clear();
            
            // 清空内存缓存
            ProofreadService.ClearCache();
            lock (_proofreadResultsLock) { _proofreadResults.Clear(); }
            
            // 重置顶部状态
            UpdateHeaderStatus("就绪", Brushes.Green);
            
            AddMessageBubble("系统", "聊天记录和缓存已清空", false);
        }

        private void InitializeProofreadConfig()
        {
            if (TxtProofreadPrompt != null && CmbProofreadMode != null)
            {
                string mode = (CmbProofreadMode.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "精准校验";
                TxtProofreadPrompt.Text = mode == "全文校验" ? GetFullTextProofreadPrompt() : GetPreciseProofreadPrompt();
            }
        }

        private string GetPreciseProofreadPrompt()
        {
            return "你是中文校对编辑。校对以下文本中的错别字、语病、标点、序号、用词、术语不一致。\n\n" +
                "规则：只标错误不改风格，每处给原文/修改/理由，不确定标[待确认]，无错误输出未发现错误。\n\n" +
                "格式：\n【第X处】类型：...\n原文：...\n修改：...\n理由：...";
        }

        private string GetFullTextProofreadPrompt()
        {
            return "你是中文校对编辑。对以下全文进行系统性校对，包括错别字、语病、标点符号、序号格式、用词准确性、术语一致性、逻辑连贯性等方面。\n\n" +
                "规则：\n1. 先给出整体质量评估和主要问题总结\n2. 再列出具体错误，每处给原文/修改/理由\n3. 不确定的地方标[待确认]\n4. 如无明显错误，输出\"全文未发现明显错误\"\n\n" +
                "格式：\n【整体评估】...\n\n【第1处】类型：...\n原文：...\n修改：...\n理由：...\n...";
        }

        private void CmbProofreadMode_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CmbProofreadMode.SelectedItem is ComboBoxItem item)
            {
                string mode = item.Content?.ToString() ?? "精准校验";
                
                // 保存校验模式到配置
                ConfigManager.CurrentConfig.ProofreadMode = mode;
                ConfigManager.SaveConfig(ConfigManager.CurrentConfig);
                
                // 根据新模式加载对应的提示词
                if (TxtProofreadPrompt != null)
                {
                    string savedPrompt = ConfigManager.GetProofreadPromptForMode(mode);
                    if (!string.IsNullOrEmpty(savedPrompt))
                    {
                        TxtProofreadPrompt.Text = savedPrompt;
                    }
                    else
                    {
                        TxtProofreadPrompt.Text = mode == "全文校验" ? GetFullTextProofreadPrompt() : GetPreciseProofreadPrompt();
                    }
                }
                
                // 显示状态提示（绿色）
                if (LblProofreadStatus != null)
                {
                    LblProofreadStatus.Text = $"状态: 已切换为{mode}";
                    LblProofreadStatus.Foreground = Brushes.Green;
                }
            }
        }

        private void BtnSaveProofreadConfig_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string prompt = TxtProofreadPrompt.Text?.Trim() ?? "";
                string currentMode = ConfigManager.CurrentConfig.ProofreadMode ?? "精准校验";
                
                // 检查提示词是否为空
                if (string.IsNullOrWhiteSpace(prompt))
                {
                    if (LblProofreadStatus != null)
                    {
                        LblProofreadStatus.Text = "状态: 提示词不能为空";
                        LblProofreadStatus.Foreground = Brushes.Red;
                    }
                    AddMessageBubble("系统", "保存失败：提示词不能为空，请输入提示词后重试", false, true);
                    return;
                }
                
                // 根据当前模式保存到对应的字段
                if (currentMode == "全文校验")
                    ConfigManager.CurrentConfig.ProofreadFullTextPrompt = prompt;
                else
                    ConfigManager.CurrentConfig.ProofreadPrecisePrompt = prompt;
                ConfigManager.SaveConfig(ConfigManager.CurrentConfig);
                
                // 更新状态提示
                if (LblProofreadStatus != null)
                {
                    LblProofreadStatus.Text = "状态: 提示词已保存";
                    LblProofreadStatus.Foreground = Brushes.Green;
                }
                AddMessageBubble("系统", $"提示词已保存（当前模式: {currentMode}）", false);
            }
            catch (Exception ex)
            {
                if (LblProofreadStatus != null)
                {
                    LblProofreadStatus.Text = $"状态: 保存失败 - {ex.Message}";
                    LblProofreadStatus.Foreground = Brushes.Red;
                }
                AddMessageBubble("错误", $"保存提示词失败: {ex.Message}", false, true);
            }
        }

        private void TabAIConfig_Click(object sender, RoutedEventArgs e)
        {
            // 显示 AI 配置面板，隐藏校验配置面板
            PanelAIConfig.Visibility = Visibility.Visible;
            PanelProofreadConfig.Visibility = Visibility.Collapsed;
            
            // 更新 Tab 样式
            TabAIConfig.BorderBrush = _primaryColor;
            TabAIConfig.Foreground = _primaryColor;
            TabAIConfig.FontWeight = FontWeights.SemiBold;
            
            TabProofreadConfig.BorderBrush = Brushes.Transparent;
            TabProofreadConfig.Foreground = _textSecondaryColor;
            TabProofreadConfig.FontWeight = FontWeights.Normal;
        }

        private void TabProofreadConfig_Click(object sender, RoutedEventArgs e)
        {
            // 显示校验配置面板，隐藏 AI 配置面板
            PanelAIConfig.Visibility = Visibility.Collapsed;
            PanelProofreadConfig.Visibility = Visibility.Visible;
            
            // 更新 Tab 样式
            TabProofreadConfig.BorderBrush = _primaryColor;
            TabProofreadConfig.Foreground = _primaryColor;
            TabProofreadConfig.FontWeight = FontWeights.SemiBold;
            
            TabAIConfig.BorderBrush = Brushes.Transparent;
            TabAIConfig.Foreground = _textSecondaryColor;
            TabAIConfig.FontWeight = FontWeights.Normal;
        }

        private void UpdateStatus(string text, Brush color)
        {
            // 更新AI配置面板状态（不再影响顶部标题栏）
            UpdateAIConfigStatus(text, color);
        }

        /// <summary>
        /// 更新AI配置面板状态
        /// </summary>
        private void UpdateAIConfigStatus(string text, Brush color)
        {
            try
            {
                if (LblAIConfigStatus != null)
                {
                    LblAIConfigStatus.Text = $"状态：{text}";
                    LblAIConfigStatus.Foreground = color;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UpdateAIConfigStatus] 错误: {ex.Message}");
            }
        }

        // 校对服务实例
        private ProofreadService _proofreadService;
        private volatile CancellationTokenSource _proofreadCts;
        private readonly List<ParagraphResult> _proofreadResults = new List<ParagraphResult>();
        private readonly object _proofreadResultsLock = new object();

        public void StartProofread()
        {
            // 使用 FireAndForget 模式，确保所有异常都被捕获
            _ = StartProofreadAsync();
        }

        private async Task StartProofreadAsync()
        {
            try
            {
                await DoStartProofreadInternal();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[StartProofread] 未捕获的异常: {ex}");
                try
                {
                    UpdateHeaderStatus("失败", Brushes.Red);
                    AddMessageBubble("错误", $"校对过程中发生错误: {ex.Message}", false, true);
                    BtnProofread.IsEnabled = true;
                }
                catch { }
            }
        }

        private async Task DoStartProofreadInternal()
        {
            if (_llmService == null)
            {
                MessageBox.Show("请先点击「保存并连接」按钮连接大模型", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string docText = _wordController.GetDocumentText();
            if (string.IsNullOrWhiteSpace(docText)) return;

            // 取消并释放之前的校对任务
            if (_proofreadCts != null)
            {
                _proofreadCts.Cancel();
                _proofreadCts.Dispose();
            }
            _proofreadCts = new CancellationTokenSource();

            // 从配置读取提示词
            var (_, systemPrompt) = ConfigManager.GetProofreadConfig();

            // 创建校对服务（并发数5）
            _proofreadService = new ProofreadService(_llmService, systemPrompt, concurrency: 5);
            _proofreadService.OnProgress += OnProofreadProgress;

            lock (_proofreadResultsLock) { _proofreadResults.Clear(); }
            
            UpdateHeaderStatus("正在校对...", Brushes.Orange);
            BtnProofread.IsEnabled = false;
            
            try
            {
                var stopwatch = Stopwatch.StartNew();
                var results = await _proofreadService.ProofreadDocumentAsync(docText, _proofreadCts.Token);
                stopwatch.Stop();
                lock (_proofreadResultsLock) { _proofreadResults.AddRange(results); }
                
                // 生成详细报告
                var report = ProofreadService.GenerateReport(results, docText.Length, stopwatch.Elapsed, _llmService?.ProviderName);
                
                // 解析所有问题项
                var allIssues = new List<ProofreadIssueItem>();
                foreach (var result in results)
                {
                    if (!string.IsNullOrWhiteSpace(result.ResultText))
                    {
                        allIssues.AddRange(ProofreadIssueParser.ParseProofreadItems(result.ResultText));
                    }
                }
                
                // 应用批注到文档，并获取处理后的项目（带位置信息）
                var processedIssues = _wordController.ApplyProofreadToDocument(allIssues, AddMessageBubble);
                
                // 在聊天框显示合并的结果（问题列表+分块详情）
                AddProofreadResultBubble("校对结果", report, processedIssues, results);
                
                // 获取缓存统计
                var (cacheCount, cacheBytes) = ProofreadService.GetCacheStats();
                
                // 更新顶部状态为完成（绿色）
                UpdateHeaderStatus($"完成 | 共 {processedIssues.Count} 处问题 | 缓存 {cacheCount} 段", Brushes.Green);
            }
            catch (OperationCanceledException)
            {
                UpdateHeaderStatus("已取消", Brushes.Gray);
                AddMessageBubble("系统", "校对已取消", false);
            }
            catch (Exception ex)
            {
                UpdateHeaderStatus("失败", Brushes.Red);
                AddMessageBubble("错误", ex.Message, false, true);
            }
            finally
            {
                BtnProofread.IsEnabled = true;
                if (_proofreadService != null)
                {
                    _proofreadService.OnProgress -= OnProofreadProgress;
                    _proofreadService.Dispose();
                    _proofreadService = null;
                }
            }
        }

        /// <summary>
        /// 更新顶部标题栏状态（线程安全）
        /// </summary>
        private void UpdateHeaderStatus(string text, Brush color)
        {
            if (Dispatcher.CheckAccess())
            {
                // 在UI线程，直接更新
                UpdateHeaderStatusInternal(text, color);
            }
            else
            {
                // 不在UI线程，使用Dispatcher
                Dispatcher.Invoke(() => UpdateHeaderStatusInternal(text, color));
            }
        }
        
        private void UpdateHeaderStatusInternal(string text, Brush color)
        {
            try
            {
                if (LblHeaderStatus != null)
                {
                    LblHeaderStatus.Text = $"状态：{text}";
                    LblHeaderStatus.Foreground = color;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UpdateHeaderStatus] 错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理校对进度更新
        /// </summary>
        private void OnProofreadProgress(object sender, ProofreadProgressArgs e)
        {
            try
            {
                if (e.IsCompleted)
                {
                    // 全部完成，在方法外更新最终状态
                    return;
                }

                // 构建状态文本
                var status = $"校对中 {e.CompletedParagraphs}/{e.TotalParagraphs}";
                if (e.CacheHitCount > 0)
                    status += $" (缓存{e.CacheHitCount})";
                
                UpdateHeaderStatus(status, Brushes.Orange);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OnProofreadProgress] 错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 生成简要统计（用于进度气泡）
        /// </summary>
        private string GenerateBriefStats(List<ParagraphResult> results)
        {
            var sb = new StringBuilder();
            
            int totalIssues = 0;
            var allCategories = new Dictionary<string, int>();
            int cachedCount = 0;
            
            foreach (var result in results)
            {
                int issues = ProofreadIssueParser.CountIssues(result.ResultText);
                totalIssues += issues;
                
                var cats = ProofreadIssueParser.CategorizeIssues(result.ResultText);
                foreach (var kv in cats)
                {
                    if (allCategories.ContainsKey(kv.Key))
                        allCategories[kv.Key] += kv.Value;
                    else
                        allCategories[kv.Key] = kv.Value;
                }
                
                if (result.IsCached)
                    cachedCount++;
            }
            
            sb.AppendLine($"共发现 {totalIssues} 处问题");
            
            if (allCategories.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("分类统计：");
                foreach (var kv in allCategories.OrderByDescending(x => x.Value).Take(5))
                {
                    sb.AppendLine($"  {kv.Key}：{kv.Value} 处");
                }
            }
            
            if (cachedCount > 0)
                sb.AppendLine($"\n缓存命中：{cachedCount} 段");
            
            return sb.ToString();
        }



        /// <summary>
        /// 添加带可点击问题列表和分块详情的校对结果气泡
        /// </summary>
        private void AddProofreadResultBubble(string reportTitle, string reportContent, List<ProofreadIssueItem> items, List<ParagraphResult> paragraphResults)
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
                Cursor = System.Windows.Input.Cursors.Hand,
                VerticalAlignment = VerticalAlignment.Center
            };
            copyButton.Click += (s, e) =>
            {
                try
                {
                    // 构建复制内容（仅报告）
                    Clipboard.SetText(reportContent);
                    copyButton.Content = "✓ 已复制";
                    Task.Delay(2000).ContinueWith(_ =>
                    {
                        Dispatcher.Invoke(() => copyButton.Content = "📋 复制");
                    });
                }
                catch { }
            };
            Grid.SetColumn(copyButton, 1);
            headerPanel.Children.Add(copyButton);
            
            mainStack.Children.Add(headerPanel);

            // 报告内容（仅统计信息，不含问题列表详情）
            var reportText = new TextBlock
            {
                Text = reportContent,
                FontSize = 13,
                Foreground = _textPrimaryColor,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 6, 0, 8)
            };
            mainStack.Children.Add(reportText);

            // 如果有问题项，在统计汇总下方直接添加可点击的问题按钮列表
            if (items.Count > 0)
            {
                // 添加分割线
                mainStack.Children.Add(new Separator 
                { 
                    Margin = new Thickness(0, 8, 0, 4),
                    Background = new SolidColorBrush(Color.FromRgb(200, 200, 200))
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

            bubbleBorder.Child = mainStack;

            var container = new Grid
            {
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(0, 0, 0, 8)
            };
            container.Children.Add(bubbleBorder);

            // 使用 ComplexMessageViewModel 承载复杂 UI
            var complexVm = new ComplexMessageViewModel 
            { 
                Content = container,
                AlignRight = false
            };
            ChatMessages.Add(complexVm);
            ChatScrollViewer.ScrollToEnd();
        }

        /// <summary>
        /// 创建问题项按钮
        /// </summary>
        private Button CreateIssueButton(ProofreadIssueItem item, int displayIndex)
        {
            var button = new Button
            {
                Background = new SolidColorBrush(Color.FromRgb(250, 250, 250)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(220, 220, 220)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(10, 8, 10, 8),
                Margin = new Thickness(0, 0, 0, 6),
                HorizontalAlignment = HorizontalAlignment.Stretch,
                HorizontalContentAlignment = HorizontalAlignment.Left,
                Cursor = System.Windows.Input.Cursors.Hand
            };

            var contentStack = new StackPanel();
            
            // 问题类型和序号行
            var headerPanel = new StackPanel { Orientation = Orientation.Horizontal };
            
            // 根据严重程度设置颜色
            Brush severityBrush = Brushes.Gray;
            if (!string.IsNullOrEmpty(item.Severity))
            {
                if (item.Severity.Contains("高") || item.Severity.Contains("严重"))
                    severityBrush = Brushes.Red;
                else if (item.Severity.Contains("中"))
                    severityBrush = new SolidColorBrush(Color.FromRgb(255, 152, 0));
                else if (item.Severity.Contains("低"))
                    severityBrush = Brushes.Green;
            }
            
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
                Background = new SolidColorBrush(Color.FromRgb(230, 230, 230))
            });
            
            // 原文
            var originalLabel = new TextBlock
            {
                Text = "原文：",
                FontSize = 10,
                FontWeight = FontWeights.SemiBold,
                Foreground = _textSecondaryColor
            };
            contentStack.Children.Add(originalLabel);
            
            var originalText = new TextBlock
            {
                Text = item.Original,
                FontSize = 11,
                Foreground = _textPrimaryColor,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(8, 0, 0, 6)
            };
            contentStack.Children.Add(originalText);
            
            // 修改
            var modifiedLabel = new TextBlock
            {
                Text = "修改：",
                FontSize = 10,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(0, 150, 0))
            };
            contentStack.Children.Add(modifiedLabel);
            
            var modifiedText = new TextBlock
            {
                Text = item.Modified,
                FontSize = 11,
                Foreground = new SolidColorBrush(Color.FromRgb(0, 120, 0)),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(8, 0, 0, 6)
            };
            contentStack.Children.Add(modifiedText);
            
            // 理由
            var reasonLabel = new TextBlock
            {
                Text = "理由：",
                FontSize = 10,
                FontWeight = FontWeights.SemiBold,
                Foreground = _textSecondaryColor
            };
            contentStack.Children.Add(reasonLabel);
            
            var reasonText = new TextBlock
            {
                Text = item.Reason,
                FontSize = 10,
                Foreground = _textSecondaryColor,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(8, 0, 0, 0)
            };
            contentStack.Children.Add(reasonText);
            
            button.Content = contentStack;
            
            // 点击事件 - 定位到文档
            button.Click += (s, e) => _wordController.NavigateToIssue(item);
            
            return button;
        }

    }
}

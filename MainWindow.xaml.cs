using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Navigation;

namespace Office2PDF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindowViewModel ViewModel { get; set; }
        private ConversionEngine _conversionEngine;
        private CancellationTokenSource _cancellationTokenSource;
        private DuplicateFileAction _duplicateFileAction = DuplicateFileAction.Skip;

        // 防止重复日志输出
        private string _lastFromFolderPath = "";
        private string _lastToFolderPath = "";

        // 版本号属性，从Assembly中提取
        public string VersionNumber
        {
            get
            {
                var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                return $"{version.Major}.{version.Minor}.{version.Build}";
            }
        }

        public MainWindow()
        {
            InitializeComponent();

            // 确保窗口大小不超出屏幕工作区域
            var workingArea = SystemParameters.WorkArea;
            if (this.Height > workingArea.Height * 0.9)
            {
                this.Height = workingArea.Height * 0.9;
            }
            if (this.Width > workingArea.Width * 0.9)
            {
                this.Width = workingArea.Width * 0.9;
            }

            ViewModel = new MainWindowViewModel();
            DataContext = ViewModel;
            LogRichTextBox.Document.Blocks.Clear();

            // 初始化转换引擎
            _conversionEngine = new ConversionEngine(ViewModel, AppendLog);

            // 初始化重复文件处理选项
            DuplicateFileActionCombo.SelectedIndex = 1; // 默认选择"跳过"

            // 注册窗口关闭事件
            this.Closing += MainWindow_Closing;

            // 初始化按钮状态
            UpdateCanStartConvert();

            // 检查命令行参数（拖拽文件夹到 exe）
            if (!string.IsNullOrEmpty(App.CommandLineFolder))
            {
                this.Loaded += (s, e) => SetFolderPaths(App.CommandLineFolder);
            }
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // 程序关闭时清理备份文件
            _conversionEngine?.ClearConversionHistory();
            // 清理网络文件的临时副本
            NetworkPathHelper.CleanupAllTempFiles();
        }

        // 统一的文件夹设置方法，防止重复日志输出
        private void SetFolderPaths(string fromPath)
        {
            var toPath = fromPath + "_PDFs";

            if (_lastFromFolderPath != fromPath || _lastToFolderPath != toPath)
            {
                ViewModel.FromRootFolderPath = fromPath;
                ViewModel.ToRootFolderPath = toPath;

                AppendLog($"来源文件夹已设置为: {fromPath}");
                AppendLog($"目标文件夹已自动设置为: {toPath}");

                _lastFromFolderPath = fromPath;
                _lastToFolderPath = toPath;
            }
        }

        // 统一的目标文件夹设置方法
        private void SetToFolderPath(string toPath)
        {
            if (_lastToFolderPath != toPath)
            {
                ViewModel.ToRootFolderPath = toPath;
                AppendLog($"目标文件夹已设置为: {toPath}");
                _lastToFolderPath = toPath;
            }
        }

        private void BrowseFromFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "选择来源文件夹";
                folderDialog.SelectedPath = ViewModel.FromRootFolderPath;
                if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    SetFolderPaths(folderDialog.SelectedPath);
                }
            }
        }

        private void BrowseToFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "选择目标文件夹";
                folderDialog.SelectedPath = ViewModel.ToRootFolderPath;
                if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    SetToFolderPath(folderDialog.SelectedPath);
                }
            }
        }

        private void OpenFromFolder_Click(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(ViewModel.FromRootFolderPath))
            {
                using (var process = new System.Diagnostics.Process())
                {
                    process.StartInfo.FileName = "explorer.exe";
                    process.StartInfo.Arguments = ViewModel.FromRootFolderPath;
                    process.Start();
                }
            }
            else
            {
                System.Windows.MessageBox.Show("来源文件夹不存在，请重新选择", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void OpenToFolder_Click(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(ViewModel.ToRootFolderPath))
            {
                using (var process = new System.Diagnostics.Process())
                {
                    process.StartInfo.FileName = "explorer.exe";
                    process.StartInfo.Arguments = ViewModel.ToRootFolderPath;
                    process.Start();
                }
            }
            else
            {
                System.Windows.MessageBox.Show("目标文件夹不存在，请重新选择", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void Folder_DragOver(object sender, System.Windows.DragEventArgs e)
        {
            if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(System.Windows.DataFormats.FileDrop);
                if (files.Length == 1 && Directory.Exists(files[0]))
                {
                    e.Effects = System.Windows.DragDropEffects.Copy;
                }
                else
                {
                    e.Effects = System.Windows.DragDropEffects.None;
                }
            }
            else
            {
                e.Effects = System.Windows.DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void FromFolder_Drop(object sender, System.Windows.DragEventArgs e)
        {
            if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(System.Windows.DataFormats.FileDrop);
                if (files.Length == 1 && Directory.Exists(files[0]))
                {
                    SetFolderPaths(files[0]);
                }
            }
        }

        private void ToFolder_Drop(object sender, System.Windows.DragEventArgs e)
        {
            if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(System.Windows.DataFormats.FileDrop);
                if (files.Length == 1 && Directory.Exists(files[0]))
                {
                    SetToFolderPath(files[0]);
                }
            }
        }

        public void UpdateCanStartConvert()
        {
            // 如果正在转换，按钮应该可用（用于取消）
            if (_cancellationTokenSource != null)
            {
                ViewModel.CanStartConvert = true;
                return;
            }

            // 检查必要条件
            bool canStart = !string.IsNullOrWhiteSpace(ViewModel.FromRootFolderPath) &&
                           !string.IsNullOrWhiteSpace(ViewModel.ToRootFolderPath) &&
                           Directory.Exists(ViewModel.FromRootFolderPath);

            ViewModel.CanStartConvert = canStart;
        }

        private async void StartConvert_Click(object sender, RoutedEventArgs e)
        {
            // 如果正在转换，则取消转换
            if (_cancellationTokenSource != null)
            {
                AppendLog("用户请求取消转换，正在停止...", LogLevel.Warning);
                ViewModel.ButtonText = "正在停止...";
                ViewModel.CanStartConvert = false;
                _cancellationTokenSource.Cancel();
                return;
            }

            // 检查是否有可撤回的更改
            if (_conversionEngine.ConversionHistoryCount > 0 && UndoButton.IsEnabled)
            {
                var result = System.Windows.MessageBox.Show(
                    $"检测到您有 {_conversionEngine.ConversionHistoryCount} 个可撤回的更改。\n\n" +
                    "开始新的转换将：\n" +
                    "• 清除所有撤回记录\n" +
                    "• 删除所有备份文件\n" +
                    "• 使撤回功能完全失效\n\n" +
                    "您确定要继续吗？如需保留撤回功能，请先完成撤回操作。",
                    "撤回功能将失效",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result == MessageBoxResult.No)
                {
                    return;
                }
            }

            // 验证路径
            if (string.IsNullOrWhiteSpace(ViewModel.FromRootFolderPath) || !Directory.Exists(ViewModel.FromRootFolderPath))
            {
                System.Windows.MessageBox.Show("请选择有效的源文件夹路径！", "路径错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(ViewModel.ToRootFolderPath))
            {
                System.Windows.MessageBox.Show("请设置目标文件夹路径！", "路径错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 验证并处理目标路径
            try
            {
                string absoluteToPath = Path.GetFullPath(ViewModel.ToRootFolderPath);
                ViewModel.ToRootFolderPath = absoluteToPath;

                if (File.Exists(absoluteToPath))
                {
                    System.Windows.MessageBox.Show("目标路径指向一个现有文件，请选择一个文件夹。", "路径错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (!Directory.Exists(absoluteToPath))
                {
                    Directory.CreateDirectory(absoluteToPath);
                    AppendLog($"创建目标文件夹：{absoluteToPath}", LogLevel.Info);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"目标文件夹路径无效或无法创建：{ex.Message}", "路径错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                // 重置引擎和状态
                _conversionEngine.Reset();
                _conversionEngine.SetDuplicateFileAction(_duplicateFileAction);
                _conversionEngine.ClearConversionHistory();
                UndoButton.IsEnabled = false;

                // 创建取消令牌
                _cancellationTokenSource = new CancellationTokenSource();
                ViewModel.ButtonText = "结束";
                UpdateCanStartConvert();

                // 验证选择的Office引擎
                if (ViewModel.UseWpsOffice)
                {
                    if (!IsWpsOfficeAvailable())
                    {
                        System.Windows.MessageBox.Show(
                            "未检测到 WPS Office 或 WPS Office 不可用。\n\n请选择：\n1. 安装 WPS Office\n2. 切换到 MS Office 引擎",
                            "WPS Office 不可用",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                        return;
                    }
                    AppendLog("提示：后台使用 WPS 引擎进行转换", LogLevel.Info);
                }
                else
                {
                    AppendLog("提示：自动选择后台引擎进行转换", LogLevel.Info);
                }

                // 定义文件类型处理器
                var fileTypeHandlers = new (string TypeName, bool IsConvert, string[] Extensions, Func<string, string[], Task> ConvertAction)[] {
                    ("Word", ViewModel.IsConvertWord, new string[] {".doc", ".docx"}, async (typeName, files) => {
                        if (ViewModel.UseWpsOffice)
                            await _conversionEngine.ConvertToPDFAsync<WpsWriterApplication>(typeName, files, _cancellationTokenSource.Token);
                        else
                            await _conversionEngine.ConvertToPDFWithAutoFallbackAsync<MSWordApplication, WpsWriterApplication>(typeName, files, _cancellationTokenSource.Token);
                    }),
                    ("Excel", ViewModel.IsConvertExcel, new string[] {".xls", ".xlsx"}, async (typeName, files) => {
                        if (ViewModel.UseWpsOffice)
                            await _conversionEngine.ConvertToPDFAsync<WpsSpreadsheetApplication>(typeName, files, _cancellationTokenSource.Token);
                        else
                            await _conversionEngine.ConvertToPDFWithAutoFallbackAsync<MSExcelApplication, WpsSpreadsheetApplication>(typeName, files, _cancellationTokenSource.Token);
                    }),
                    ("PPT", ViewModel.IsConvertPPT, new string[] {".ppt", ".pptx"}, async (typeName, files) => {
                        if (ViewModel.UseWpsOffice)
                            await _conversionEngine.ConvertToPDFAsync<WpsPresentationApplication>(typeName, files, _cancellationTokenSource.Token);
                        else
                            await _conversionEngine.ConvertToPDFWithAutoFallbackAsync<MSPowerPointApplication, WpsPresentationApplication>(typeName, files, _cancellationTokenSource.Token);
                    })
                };

                AppendLog($"==============开始转换==============");

                // 预扫描：检测文件名冲突
                _conversionEngine.PreScanFilesForConflicts(fileTypeHandlers);

                // 检测网络路径
                bool hasNetworkPath = NetworkPathHelper.IsNetworkPath(ViewModel.FromRootFolderPath) ||
                                     NetworkPathHelper.IsNetworkPath(ViewModel.ToRootFolderPath);
                if (hasNetworkPath)
                {
                    AppendLog($"💡 检测到网络路径，启用本地临时处理策略（所有文件类型）", LogLevel.Info);
                }

                // 统计文件数量
                int totalWordCount = 0;
                int totalExcelCount = 0;
                int totalPptCount = 0;

                // 处理各种文件类型
                foreach (var (typeName, isConvert, extensions, convertAction) in fileTypeHandlers)
                {
                    if (_cancellationTokenSource.Token.IsCancellationRequested)
                    {
                        AppendLog("转换已被取消", LogLevel.Warning);
                        break;
                    }

                    try
                    {
                        if (isConvert)
                        {
                            AppendLog($"{typeName} 转换开始", LogLevel.Info);

                            var searchOption = ViewModel.IsConvertChildrenFolder ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                            var files = Directory.EnumerateFiles(ViewModel.FromRootFolderPath, "*.*", searchOption)
                             .Where(file => extensions.Any(ext => file.EndsWith(ext, StringComparison.OrdinalIgnoreCase))).ToArray();

                            if (!files.Any())
                            {
                                AppendLog($"无 {typeName} 文件", LogLevel.Warning);
                            }
                            else
                            {
                                // 统计文件数量
                                if (typeName == "Word") totalWordCount = files.Length;
                                else if (typeName == "Excel") totalExcelCount = files.Length;
                                else if (typeName == "PPT") totalPptCount = files.Length;

                                await Task.Run(async () => await convertAction(typeName, files), _cancellationTokenSource.Token);
                            }
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        AppendLog($"{typeName} 转换被取消", LogLevel.Warning);
                        break;
                    }
                    catch (Exception ex)
                    {
                        AppendLog($"{typeName} 转换错误：{ex.Message}", LogLevel.Error);
                    }
                    finally
                    {
                        if (isConvert)
                        {
                            AppendLog($"{typeName} 转换结束", LogLevel.Info);
                        }
                    }
                }

                // 设置总文件数
                _conversionEngine.SetTotalFilesCount(totalWordCount, totalExcelCount, totalPptCount);
            }
            catch (OperationCanceledException)
            {
                AppendLog("转换被用户取消", LogLevel.Warning);
            }
            catch (Exception ex)
            {
                AppendLog($"转换错误：{ex.Message}", LogLevel.Error);
            }
            finally
            {
                // 显示转换结果汇总
                _conversionEngine.ShowConversionSummary();

                // 处理文件删除
                if (_conversionEngine.SuccessfullyConvertedFiles.Count > 0 && !_cancellationTokenSource.Token.IsCancellationRequested)
                {
                    await DeleteOriginalFilesAsync();
                }

                // 重置状态
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
                ViewModel.ButtonText = "开始";
                UpdateCanStartConvert();

                // 如果有转换操作，启用撤回按钮
                if (_conversionEngine.ConversionHistoryCount > 0)
                {
                    UndoButton.IsEnabled = true;
                    AppendLog($"转换完成！如需撤回所有更改，请点击\"撤回更改\"按钮", LogLevel.Info);
                }
            }
        }

        private async Task DeleteOriginalFilesAsync()
        {
            if (!ViewModel.IsDeleteOriginalFiles)
                return;

            var (deletedCount, failedFiles) = await _conversionEngine.DeleteOriginalFilesAsync();
            // 日志已在引擎中输出
        }

        public void AppendLog(string message, LogLevel level = LogLevel.Trace)
        {
            Dispatcher.Invoke(() =>
            {
                var paragraph = new System.Windows.Documents.Paragraph();
                paragraph.Inlines.Add(new Run($"[{DateTime.Now:HH:mm:ss}] ") { Foreground = Brushes.Gray });

                Brush color = Brushes.Black;
                switch (level)
                {
                    case LogLevel.Trace: color = Brushes.Gray; break;
                    case LogLevel.Info: color = Brushes.Green; break;
                    case LogLevel.Warning: color = Brushes.Orange; break;
                    case LogLevel.Error: color = Brushes.Red; break;
                }

                paragraph.Inlines.Add(new Run(message) { Foreground = color });
                LogRichTextBox.Document.Blocks.Add(paragraph);
                LogRichTextBox.ScrollToEnd();
            });
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = e.Uri.AbsoluteUri,
                UseShellExecute = true
            });
            e.Handled = true;
        }

        private void SoftwareHomepage_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Hyperlink hyperlink && hyperlink.ContextMenu != null)
            {
                hyperlink.ContextMenu.PlacementTarget = null;
                hyperlink.ContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.MousePoint;
                hyperlink.ContextMenu.IsOpen = true;
            }
        }

        private void Homepage_Current_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.MenuItem menuItem && menuItem.Tag is string url)
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
            }
        }

        private void Homepage_Source1_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.MenuItem menuItem && menuItem.Tag is string url)
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
            }
        }

        private void Homepage_Source2_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.MenuItem menuItem && menuItem.Tag is string url)
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
            }
        }

        private void AboutButton_Click(object sender, RoutedEventArgs e)
        {
            var aboutWindow = new AboutWindow();
            aboutWindow.Owner = this;
            aboutWindow.ShowDialog();
        }

        private void HelpButton_Click(object sender, RoutedEventArgs e)
        {
            EngineHelpPopup.IsOpen = true;
        }

        private void DuplicateFileActionCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DuplicateFileActionCombo.SelectedItem is ComboBoxItem selectedItem)
            {
                var tag = selectedItem.Tag?.ToString();
                if (Enum.TryParse(tag, out DuplicateFileAction action))
                {
                    _duplicateFileAction = action;
                    _conversionEngine?.SetDuplicateFileAction(action);

                    switch (action)
                    {
                        case DuplicateFileAction.Skip:
                            AppendLog("⏭️ 目标文件已存在时：自动跳过（推荐）", LogLevel.Info);
                            break;
                        case DuplicateFileAction.Overwrite:
                            AppendLog("🔄 目标文件已存在时：自动覆盖", LogLevel.Warning);
                            break;
                        case DuplicateFileAction.Rename:
                            AppendLog("📝 目标文件已存在时：自动重命名", LogLevel.Info);
                            break;
                    }
                }
            }
        }

        private bool IsWpsOfficeAvailable()
        {
            try
            {
                var wpsProgIds = new[] { "KWps.Application", "KET.Application", "KWPP.Application" };

                foreach (var progId in wpsProgIds)
                {
                    try
                    {
                        var type = Type.GetTypeFromProgID(progId, throwOnError: false);
                        if (type == null) continue;
                        var app = Activator.CreateInstance(type);
                        try
                        {
                            if (app != null)
                            {
                                if (Marshal.IsComObject(app))
                                    Marshal.ReleaseComObject(app);
                                return true;
                            }
                        }
                        finally
                        {
                            app = null;
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                    }
                    catch
                    {
                        // 继续检查下一个
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        private async void UndoChanges_Click(object sender, RoutedEventArgs e)
        {
            if (_conversionEngine.ConversionHistoryCount == 0)
            {
                AppendLog("没有可撤回的操作", LogLevel.Info);
                return;
            }

            var result = System.Windows.MessageBox.Show(
                $"确定要撤回本次转换的所有更改吗？\n\n将会删除 {_conversionEngine.ConversionHistoryCount} 个转换生成的文件。\n此操作不可逆！",
                "确认撤回",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                UndoButton.IsEnabled = false;
                var (undoneCount, failedCount) = await _conversionEngine.PerformUndoAsync();
                // 日志已在引擎中输出
            }
        }
    }
}

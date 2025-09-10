using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Windows.Documents;
using System.Windows.Media;
using System.Diagnostics;
using System.Windows.Navigation;
using System.Windows.Forms;
using System.Linq;
using System.Globalization;
using System.Windows.Data;
using System.Runtime.InteropServices;
using System.Threading;

namespace Office2PDF
{
    // 布尔值反转转换器
    public class InverseBooleanConverter : IValueConverter
    {
        public static readonly InverseBooleanConverter Instance = new InverseBooleanConverter();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value is bool b ? !b : false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value is bool b ? !b : false;
        }
    }

    // 重复文件处理策略
    public enum DuplicateFileAction
    {
        Skip,       // 跳过
        Overwrite,  // 覆盖
        Rename      // 智能重命名
    }

    public class FileHandleResult
    {
        public string FilePath { get; set; }
        public DuplicateFileAction Action { get; set; }
        public bool IsOriginalFile { get; set; }
    }

    // 撤回功能的操作记录
    public enum OperationType
    {
        CreateFile,      // 创建新文件
        OverwriteFile,   // 覆盖现有文件
        CreateDirectory, // 创建新目录
        DeleteFile       // 删除文件（删除原文件功能）
    }

    public class ConversionOperation
    {
        public OperationType Type { get; set; }
        public string FilePath { get; set; }          // 操作的文件路径
        public string BackupPath { get; set; }        // 备份文件路径(覆盖时使用)
        public DateTime Timestamp { get; set; }       // 操作时间
        public string SourceFile { get; set; }        // 源文件路径
    }

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindowViewModel ViewModel { get; set; }

        // 版本号属性，从Assembly中提取
        public string VersionNumber
        {
            get
            {
                var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                return $"{version.Major}.{version.Minor}.{version.Build}";
            }
        }

        private List<string> _successfullyConvertedFiles = new List<string>();
        private List<string> _failedFiles = new List<string>();
        private List<string> _skippedFiles = new List<string>();
        private List<string> _overwrittenFiles = new List<string>();
        private List<string> _renamedFiles = new List<string>();
        private int _totalFilesCount = 0;
        private bool _wasCancelled = false;
        private CancellationTokenSource _cancellationTokenSource;

        // 按文件类型统计
        private int _successfulWordCount = 0;
        private int _successfulExcelCount = 0;
        private int _successfulPptCount = 0;
        private int _totalWordCount = 0;
        private int _totalExcelCount = 0;
        private int _totalPptCount = 0;

        // 重复文件处理策略 - 默认跳过，最安全的选择
        private DuplicateFileAction _duplicateFileAction = DuplicateFileAction.Skip;

        // 防止重复日志输出
        private string _lastFromFolderPath = "";
        private string _lastToFolderPath = "";

        // 撤回功能相关
        private List<ConversionOperation> _conversionHistory = new List<ConversionOperation>();
        private readonly string _sessionId = DateTime.Now.ToString("yyyyMMdd_HHmmss_") + Guid.NewGuid().ToString("N").Substring(0, 8); // 会话唯一标识

        public enum LogLevel
        {
            Trace,
            Info,
            Warning,
            Error
        }

        public MainWindow()
        {
            InitializeComponent();

            // 设置DataContext以支持版本号绑定
            this.DataContext = this;

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

            // 初始化重复文件处理选项
            DuplicateFileActionCombo.SelectedIndex = 1; // 默认选择"跳过"

            // 注册窗口关闭事件，确保备份文件被清理
            this.Closing += MainWindow_Closing;

            // 初始化按钮状态
            UpdateCanStartConvert();
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // 程序关闭时清理当前会话的备份文件
            CleanupCurrentSessionBackups();
        }

        // 统一的文件夹设置方法，防止重复日志输出
        private void SetFolderPaths(string fromPath)
        {
            var toPath = fromPath + "_PDFs";

            // 只有当路径真正发生变化时才输出日志
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

        // 统一的目标文件夹设置方法，防止重复日志输出
        private void SetToFolderPath(string toPath)
        {
            // 只有当路径真正发生变化时才输出日志
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
            // 检查是否是文件夹拖拽
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

        /// <summary>
        /// 更新开始按钮的可用状态
        /// </summary>
        public void UpdateCanStartConvert()
        {
            // 如果正在转换，按钮应该可用（用于取消）
            if (_cancellationTokenSource != null)
            {
                ViewModel.CanStartConvert = true;
                return;
            }

            // 检查必要条件：
            // 1. 源路径必须存在且有效（因为要读取文件）
            // 2. 目标路径只需要不为空（可以创建不存在的目录）
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

                // 立即更改按钮状态，提供即时反馈
                ViewModel.ButtonText = "正在停止...";
                ViewModel.CanStartConvert = false;

                _wasCancelled = true;
                _cancellationTokenSource.Cancel();
                return;
            }

            // 检查是否有可撤回的更改，提醒用户新转换会使撤回功能失效
            if (_conversionHistory.Count > 0 && UndoButton.IsEnabled)
            {
                var result = System.Windows.MessageBox.Show(
                    $"检测到您有 {_conversionHistory.Count} 个可撤回的更改。\n\n" +
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
                    return; // 用户选择不继续，保留撤回功能
                }
            }

            // 验证必要的路径
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
                // 将路径转换为绝对路径，这能暴露无效字符或格式问题
                string absoluteToPath = Path.GetFullPath(ViewModel.ToRootFolderPath);
                ViewModel.ToRootFolderPath = absoluteToPath; // 更新UI

                // 检查路径是否指向一个文件
                if (File.Exists(absoluteToPath))
                {
                    System.Windows.MessageBox.Show("目标路径指向一个现有文件，请选择一个文件夹。", "路径错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // 如果目标文件夹不存在，则创建它
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
                // 重置统计变量
                _successfullyConvertedFiles.Clear();
                _failedFiles.Clear();
                _skippedFiles.Clear();
                _overwrittenFiles.Clear();
                _renamedFiles.Clear();
                _totalFilesCount = 0;
                _wasCancelled = false;
                _successfulWordCount = 0;
                _successfulExcelCount = 0;
                _successfulPptCount = 0;
                _totalWordCount = 0;
                _totalExcelCount = 0;
                _totalPptCount = 0;

                // 清除撤回历史记录、清理备份文件并禁用撤回按钮
                CleanupBackupFiles(); // 先清理备份文件
                _conversionHistory.Clear(); // 再清除历史记录
                UndoButton.IsEnabled = false;

                // 创建取消令牌
                _cancellationTokenSource = new CancellationTokenSource();
                ViewModel.ButtonText = "结束";
                UpdateCanStartConvert();  // 更新按钮状态（转换时应该可用以便取消）

                // 验证选择的Office引擎是否可用
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
                    AppendLog("使用 WPS Office 引擎进行转换", LogLevel.Info);
                }
                else
                {
                    AppendLog("使用 MS Office 引擎进行转换", LogLevel.Info);
                }

                var fileTypeHandlers = new (string TypeName, bool IsConvert, string[] Extensions, Action<string, string[]> ConvertAction)[] {
                    ("Word", ViewModel.IsConvertWord, new string[] {".doc", ".docx"}, (typeName, files) => {
                        _totalWordCount += files.Length;
                        if (ViewModel.UseWpsOffice)
                            ConvertToPDF<WpsWriterApplication>(typeName, files);
                        else
                            ConvertToPDF<WordApplication>(typeName, files);
                    }),
                    ("Excel", ViewModel.IsConvertExcel, new string[] {".xls", ".xlsx"}, (typeName, files) => {
                        _totalExcelCount += files.Length;
                        if (ViewModel.UseWpsOffice)
                            ConvertToPDF<WpsSpreadsheetApplication>(typeName, files);
                        else
                            ConvertToPDF<ExcelApplication>(typeName, files);
                    }),
                    ("PPT", ViewModel.IsConvertPPT, new string[] {".ppt", ".pptx"}, (typeName, files) => {
                        _totalPptCount += files.Length;
                        if (ViewModel.UseWpsOffice)
                            ConvertToPDF<WpsPresentationApplication>(typeName, files);
                        else
                            ConvertToPDF<PowerPointApplication>(typeName, files);
                    })
                };

                AppendLog($"==============开始转换==============");

                foreach (var (typeName, isConvert, extensions, convertAction) in fileTypeHandlers)
                {
                    // 检查是否被取消
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
                                _totalFilesCount += files.Length;  // 累计文件总数
                                await Task.Run(() => convertAction(typeName, files), _cancellationTokenSource.Token);
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
                AppendLog($"==============转换结束==============");

                // 显示转换结果汇总
                ShowConversionSummary();

                // 处理文件删除
                if (_successfullyConvertedFiles.Count > 0 && !_cancellationTokenSource.Token.IsCancellationRequested)
                {
                    await DeleteOriginalFilesAsync();
                }

                // 重置状态
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
                ViewModel.ButtonText = "开始";
                UpdateCanStartConvert();  // 根据路径有效性更新按钮状态

                // 如果有转换操作（无论是否被取消），启用撤回按钮
                if (_conversionHistory.Count > 0)
                {
                    UndoButton.IsEnabled = true;
                    if (_wasCancelled)
                    {
                        AppendLog($"转换已取消！已转换 {_conversionHistory.Count} 个文件，如需撤回这些更改，请点击\"撤回更改\"按钮", LogLevel.Warning);
                    }
                    else
                    {
                        AppendLog($"转换完成！如需撤回所有更改，请点击\"撤回更改\"按钮", LogLevel.Info);
                    }
                }
            }
        }

        public void ConvertToPDF<T>(string typeName, string[] fromFilePaths) where T : IOfficeApplication, new()
        {
            using (var application = new T())
            {
                try
                {
                    AppendLog($"{typeName} 打开进程中...");
                    var numberFormat = $"D{fromFilePaths.Length.ToString().Length}";
                    for (int i = 0; i < fromFilePaths.Length; i++)
                    {
                        // 优先检查取消令牌
                        if (_cancellationTokenSource?.Token.IsCancellationRequested == true)
                        {
                            _wasCancelled = true;
                            AppendLog($"{typeName} 转换已被用户取消", LogLevel.Warning);
                            break;
                        }

                        var index = i + 1;  // 从1开始计数，符合人类习惯
                        var fromFilePath = fromFilePaths[i];
                        try
                        {
                            if (application is WordApplication wordApp)
                            {
                                wordApp.IsPrintRevisions = ViewModel.IsPrintRevisionsInWord;
                            }
                            else if (application is WpsWriterApplication wpsWordApp)
                            {
                                wpsWordApp.IsPrintRevisions = ViewModel.IsPrintRevisionsInWord;
                            }
                            else if (application is ExcelApplication excelApp)
                            {
                                excelApp.IsConvertOneSheetOnePDF = ViewModel.IsConvertOneSheetOnePDFInExcel;
                            }
                            else if (application is WpsSpreadsheetApplication wpsExcelApp)
                            {
                                wpsExcelApp.IsConvertOneSheetOnePDF = ViewModel.IsConvertOneSheetOnePDFInExcel;
                            }

                            // 在打开文档前再次检查取消
                            if (_cancellationTokenSource?.Token.IsCancellationRequested == true)
                            {
                                _wasCancelled = true;
                                AppendLog($"{typeName} 转换已被用户取消", LogLevel.Warning);
                                break;
                            }

                            application.OpenDocument(fromFilePath);
                            var toFilePath = GetToFilePath(ViewModel.FromRootFolderPath, ViewModel.ToRootFolderPath, fromFilePath, Path.GetFileName(fromFilePath));

                            // 处理重复文件
                            var handleResult = HandleDuplicateFile(toFilePath);
                            if (handleResult.FilePath == null)
                            {
                                // 用户选择跳过此文件
                                _skippedFiles.Add(fromFilePath);
                                AppendLog($"（{index.ToString(numberFormat)}）{typeName} 已跳过: {Path.GetFileName(toFilePath)} (目标文件已存在)");
                                continue;
                            }

                            // 在保存前再次检查取消
                            if (_cancellationTokenSource?.Token.IsCancellationRequested == true)
                            {
                                _wasCancelled = true;
                                AppendLog($"{typeName} 转换已被用户取消", LogLevel.Warning);
                                break;
                            }

                            // 记录处理类型
                            if (!handleResult.IsOriginalFile)
                            {
                                switch (handleResult.Action)
                                {
                                    case DuplicateFileAction.Overwrite:
                                        _overwrittenFiles.Add(fromFilePath);
                                        break;
                                    case DuplicateFileAction.Rename:
                                        _renamedFiles.Add(fromFilePath);
                                        break;
                                }
                            }

                            // 记录即将进行的操作用于撤回功能
                            RecordConversionOperation(handleResult.FilePath, fromFilePath, handleResult.Action);

                            // 执行转换
                            application.SaveAsPDF(handleResult.FilePath);

                            // 检查Excel Sheet分离模式是否生成了多个文件
                            List<string> actualGeneratedFiles = new List<string>();
                            bool isExcelApplication = application is ExcelApplication || application is WpsSpreadsheetApplication;
                            if (isExcelApplication && ViewModel.IsConvertOneSheetOnePDFInExcel)
                            {
                                // Excel Sheet分离模式：查找实际生成的文件
                                var directory = Path.GetDirectoryName(handleResult.FilePath);
                                var baseFileName = Path.GetFileNameWithoutExtension(handleResult.FilePath);
                                var extension = Path.GetExtension(handleResult.FilePath);

                                if (Directory.Exists(directory))
                                {
                                    // 查找两种可能的Sheet文件格式
                                    // MS Office格式: filename_sheet1.pdf, filename_sheet2.pdf
                                    var msOfficePattern = $"{baseFileName}_sheet*.pdf";
                                    var msOfficeFiles = Directory.GetFiles(directory, msOfficePattern);

                                    // WPS格式: filename_SheetName.pdf (实际Sheet名称)
                                    var wpsPattern = $"{baseFileName}_*.pdf";
                                    var wpsFiles = Directory.GetFiles(directory, wpsPattern);

                                    // 合并所有找到的文件，并排除非Sheet文件
                                    var allSheetFiles = msOfficeFiles.Concat(wpsFiles)
                                        .Where(f => !f.Equals(handleResult.FilePath, StringComparison.OrdinalIgnoreCase))
                                        .Distinct()
                                        .ToList();

                                    if (allSheetFiles.Count > 0)
                                    {
                                        actualGeneratedFiles.AddRange(allSheetFiles);

                                        // 删除原始操作记录，重新记录实际生成的文件
                                        if (_conversionHistory.Count > 0)
                                        {
                                            _conversionHistory.RemoveAt(_conversionHistory.Count - 1);
                                        }

                                        // 为每个Sheet文件记录操作
                                        foreach (var sheetFile in allSheetFiles)
                                        {
                                            RecordConversionOperation(sheetFile, fromFilePath, DuplicateFileAction.Rename);
                                        }
                                    }
                                    else
                                    {
                                        actualGeneratedFiles.Add(handleResult.FilePath);
                                    }
                                }
                                else
                                {
                                    actualGeneratedFiles.Add(handleResult.FilePath);
                                }
                            }
                            else
                            {
                                actualGeneratedFiles.Add(handleResult.FilePath);
                            }

                            // 输出转换结果日志
                            if (actualGeneratedFiles.Count == 1)
                            {
                                var generatedFile = actualGeneratedFiles[0];
                                var logMessage = generatedFile == toFilePath
                                    ? $"（{index.ToString(numberFormat)}）{typeName} 转换成功: {GetRelativePath(ViewModel.ToRootFolderPath, generatedFile)}"
                                    : $"（{index.ToString(numberFormat)}）{typeName} 转换成功: {GetRelativePath(ViewModel.ToRootFolderPath, generatedFile)} (已重命名)";
                                AppendLog(logMessage);
                            }
                            else
                            {
                                // Excel Sheet分离模式生成了多个文件
                                AppendLog($"（{index.ToString(numberFormat)}）{typeName} 转换成功，生成 {actualGeneratedFiles.Count} 个Sheet PDF:");
                                foreach (var file in actualGeneratedFiles)
                                {
                                    AppendLog($"    • {GetRelativePath(ViewModel.ToRootFolderPath, file)}");
                                }
                            }

                            // 统计成功转换的文件类型
                            IncrementSuccessCount(fromFilePath);

                            // 如果选择了删除原文件，则将文件路径添加到待删除列表
                            if (ViewModel.IsDeleteOriginalFiles)
                            {
                                _successfullyConvertedFiles.Add(fromFilePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            _failedFiles.Add(fromFilePath);  // 记录失败的文件
                            AppendLog($"（{index.ToString(numberFormat)}）{typeName} 转换出错：{fromFilePath} {ex.Message}", LogLevel.Error);
                        }
                        finally
                        {
                            try
                            {
                                application.CloseDocument();
                            }
                            catch (Exception)
                            {
                                // 忽略关闭文档时的异常，特别是在取消操作时
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    AppendLog(e.Message, LogLevel.Error);
                }
                finally
                {
                    AppendLog($"{typeName} 所有文件已转换完毕，关闭进程中...");
                }
            }
        }

        private string GetToFilePath(string fromRootFolderPath, string toFolderRootPath, string fromFilePath, string toFileName)
        {
            var relativePath = ".";
            if (ViewModel.IsKeepFolderStructure)
            {
                var fromFileDir = Path.GetDirectoryName(fromFilePath);
                if (!string.IsNullOrEmpty(fromFileDir))
                {
                    relativePath = GetRelativePath(fromRootFolderPath, fromFileDir);
                }
            }
            var toFolderPath = Path.Combine(toFolderRootPath, relativePath);

            // 记录新创建的目录用于撤回
            if (!Directory.Exists(toFolderPath))
            {
                Directory.CreateDirectory(toFolderPath);
                RecordDirectoryCreation(toFolderPath);
            }

            return Path.Combine(toFolderPath, Path.ChangeExtension(toFileName, ".pdf"));
        }

        private string GetRelativePath(string fromPath, string toPath)
        {
            if (string.IsNullOrEmpty(fromPath)) throw new ArgumentNullException(nameof(fromPath));
            if (string.IsNullOrEmpty(toPath)) throw new ArgumentNullException(nameof(toPath));

            Uri fromUri = new Uri(AppendDirectorySeparatorChar(fromPath));
            Uri toUri = new Uri(AppendDirectorySeparatorChar(toPath));

            if (fromUri.Scheme != toUri.Scheme) { return toPath; } // path can't be made relative.

            Uri relativeUri = fromUri.MakeRelativeUri(toUri);
            string relativePath = Uri.UnescapeDataString(relativeUri.ToString());

            if (toUri.Scheme.Equals("file", StringComparison.InvariantCultureIgnoreCase))
            {
                relativePath = relativePath.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
            }

            return relativePath;
        }

        private string AppendDirectorySeparatorChar(string path)
        {
            // Append a slash only if the path is a directory and does not have a slash.
            if (!Path.HasExtension(path) &&
                !path.EndsWith(Path.DirectorySeparatorChar.ToString()))
            {
                return path + Path.DirectorySeparatorChar;
            }

            return path;
        }

        public void AppendLog(string message, LogLevel level = LogLevel.Trace)
        {
            Dispatcher.Invoke(() => // 在 UI 线程中更新日志
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
                // 设置菜单的位置为当前鼠标位置
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

        private void DuplicateFileActionCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DuplicateFileActionCombo.SelectedItem is ComboBoxItem selectedItem)
            {
                var tag = selectedItem.Tag?.ToString();
                if (Enum.TryParse(tag, out DuplicateFileAction action))
                {
                    _duplicateFileAction = action;

                    // 显示设置信息
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

        private void IncrementSuccessCount(string filePath)
        {
            var extension = Path.GetExtension(filePath).ToLower();
            switch (extension)
            {
                case ".doc":
                case ".docx":
                    _successfulWordCount++;
                    break;
                case ".xls":
                case ".xlsx":
                    _successfulExcelCount++;
                    break;
                case ".ppt":
                case ".pptx":
                    _successfulPptCount++;
                    break;
            }
        }

        private FileHandleResult HandleDuplicateFile(string originalPath)
        {
            // 输入参数验证
            if (string.IsNullOrEmpty(originalPath))
            {
                throw new ArgumentException("文件路径不能为空", nameof(originalPath));
            }

            if (!File.Exists(originalPath))
                return new FileHandleResult
                {
                    FilePath = originalPath,
                    Action = DuplicateFileAction.Rename, // 修正：文件不存在时应该是创建新文件
                    IsOriginalFile = true
                };

            switch (_duplicateFileAction)
            {
                case DuplicateFileAction.Skip:
                    return new FileHandleResult
                    {
                        FilePath = null,
                        Action = DuplicateFileAction.Skip,
                        IsOriginalFile = false
                    };

                case DuplicateFileAction.Overwrite:
                    return new FileHandleResult
                    {
                        FilePath = originalPath,
                        Action = DuplicateFileAction.Overwrite,
                        IsOriginalFile = false
                    };

                case DuplicateFileAction.Rename:
                default:
                    // 默认使用智能重命名策略，避免转换过程中的弹窗干扰
                    // 这样用户可以通过"结束"按钮正常取消操作，不会被弹窗阻塞
                    return new FileHandleResult
                    {
                        FilePath = GetUniqueFilePath(originalPath),
                        Action = DuplicateFileAction.Rename,
                        IsOriginalFile = false
                    };
            }
        }

        private string GetUniqueFilePath(string originalPath)
        {
            // 输入验证
            if (string.IsNullOrEmpty(originalPath))
            {
                throw new ArgumentException("文件路径不能为空", nameof(originalPath));
            }

            var directory = Path.GetDirectoryName(originalPath);
            var fileNameWithoutExt = Path.GetFileNameWithoutExtension(originalPath);
            var extension = Path.GetExtension(originalPath);

            // 验证目录是否存在
            if (!Directory.Exists(directory))
            {
                throw new DirectoryNotFoundException($"目录不存在: {directory}");
            }

            int counter = 1;
            string newPath;
            const int maxAttempts = 9999; // 防止无限循环

            do
            {
                var newFileName = $"{fileNameWithoutExt} ({counter}){extension}";
                newPath = Path.Combine(directory, newFileName);
                counter++;

                if (counter > maxAttempts)
                {
                    // 如果达到最大尝试次数，添加时间戳
                    var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    newFileName = $"{fileNameWithoutExt}_{timestamp}{extension}";
                    newPath = Path.Combine(directory, newFileName);
                    break;
                }
            }
            while (File.Exists(newPath));

            return newPath;
        }

        private int CountFilesByExtension(List<string> files, string[] extensions)
        {
            return files.Count(file =>
            {
                var ext = Path.GetExtension(file).ToLower();
                return extensions.Contains(ext);
            });
        }

        private void ShowConversionSummary()
        {
            var failedCount = _failedFiles.Count;
            var skippedCount = _skippedFiles.Count;
            var overwrittenCount = _overwrittenFiles.Count;
            var renamedCount = _renamedFiles.Count;

            // 实际成功转换的数量（基于成功计数器，而不是减法计算）
            var actualSuccessCount = _successfulWordCount + _successfulExcelCount + _successfulPptCount;
            var successCount = _wasCancelled ? actualSuccessCount : _totalFilesCount - failedCount - skippedCount;

            AppendLog($"📊 ============== 转换结果汇总 ==============");

            // 显示总文件数及各类型分布
            var totalDetails = new List<string>();
            if (_totalWordCount > 0) totalDetails.Add($"📄Word {_totalWordCount}");
            if (_totalExcelCount > 0) totalDetails.Add($"📈Excel {_totalExcelCount}");
            if (_totalPptCount > 0) totalDetails.Add($"📽️PPT {_totalPptCount}");

            var totalDetailStr = totalDetails.Count > 0 ? $" | {string.Join(" + ", totalDetails)}" : "";
            AppendLog($"📁 总共文件数：{_totalFilesCount} 个{totalDetailStr}");

            // 显示成功数及各类型分布
            var successDetails = new List<string>();
            if (_successfulWordCount > 0) successDetails.Add($"📄Word {_successfulWordCount}");
            if (_successfulExcelCount > 0) successDetails.Add($"📈Excel {_successfulExcelCount}");
            if (_successfulPptCount > 0) successDetails.Add($"📽️PPT {_successfulPptCount}");

            var successDetailStr = successDetails.Count > 0 ? $" | {string.Join(" + ", successDetails)}" : "";
            AppendLog($"✅ 转换成功：{successCount} 个{successDetailStr}");

            // 显示跳过文件详情 - 按文件类型分类
            if (skippedCount > 0)
            {
                var skippedWordCount = CountFilesByExtension(_skippedFiles, new[] { ".doc", ".docx" });
                var skippedExcelCount = CountFilesByExtension(_skippedFiles, new[] { ".xls", ".xlsx" });
                var skippedPptCount = CountFilesByExtension(_skippedFiles, new[] { ".ppt", ".pptx" });

                var skippedDetails = new List<string>();
                if (skippedWordCount > 0) skippedDetails.Add($"📄Word {skippedWordCount}");
                if (skippedExcelCount > 0) skippedDetails.Add($"📈Excel {skippedExcelCount}");
                if (skippedPptCount > 0) skippedDetails.Add($"📽️PPT {skippedPptCount}");

                var skippedDetailStr = skippedDetails.Count > 0 ? $" | {string.Join(" + ", skippedDetails)}" : "";
                AppendLog($"⏭️ 跳过文件：{skippedCount} 个 (目标文件已存在){skippedDetailStr}", LogLevel.Warning);
            }

            // 显示覆盖文件详情 - 按文件类型分类
            if (overwrittenCount > 0)
            {
                var overwrittenWordCount = CountFilesByExtension(_overwrittenFiles, new[] { ".doc", ".docx" });
                var overwrittenExcelCount = CountFilesByExtension(_overwrittenFiles, new[] { ".xls", ".xlsx" });
                var overwrittenPptCount = CountFilesByExtension(_overwrittenFiles, new[] { ".ppt", ".pptx" });

                var overwrittenDetails = new List<string>();
                if (overwrittenWordCount > 0) overwrittenDetails.Add($"📄Word {overwrittenWordCount}");
                if (overwrittenExcelCount > 0) overwrittenDetails.Add($"📈Excel {overwrittenExcelCount}");
                if (overwrittenPptCount > 0) overwrittenDetails.Add($"📽️PPT {overwrittenPptCount}");

                var overwrittenDetailStr = overwrittenDetails.Count > 0 ? $" | {string.Join(" + ", overwrittenDetails)}" : "";
                AppendLog($"🔄 覆盖文件：{overwrittenCount} 个 (已覆盖同名目标文件){overwrittenDetailStr}", LogLevel.Warning);
            }

            // 显示重命名文件详情 - 按文件类型分类
            if (renamedCount > 0)
            {
                var renamedWordCount = CountFilesByExtension(_renamedFiles, new[] { ".doc", ".docx" });
                var renamedExcelCount = CountFilesByExtension(_renamedFiles, new[] { ".xls", ".xlsx" });
                var renamedPptCount = CountFilesByExtension(_renamedFiles, new[] { ".ppt", ".pptx" });

                var renamedDetails = new List<string>();
                if (renamedWordCount > 0) renamedDetails.Add($"📄Word {renamedWordCount}");
                if (renamedExcelCount > 0) renamedDetails.Add($"📈Excel {renamedExcelCount}");
                if (renamedPptCount > 0) renamedDetails.Add($"📽️PPT {renamedPptCount}");

                var renamedDetailStr = renamedDetails.Count > 0 ? $" | {string.Join(" + ", renamedDetails)}" : "";
                AppendLog($"📝 重命名文件：{renamedCount} 个 (已自动重命名){renamedDetailStr}", LogLevel.Info);
            }

            if (failedCount > 0)
            {
                AppendLog($"❌ 转换失败：{failedCount} 个", LogLevel.Error);
                AppendLog($"💥 失败文件列表：", LogLevel.Error);
                for (int i = 0; i < _failedFiles.Count; i++)
                {
                    var fileName = Path.GetFileName(_failedFiles[i]);
                    var relativePath = GetRelativePath(ViewModel.FromRootFolderPath, _failedFiles[i]);
                    AppendLog($"   {i + 1}. {relativePath}", LogLevel.Error);
                }
            }

            // 根据转换结果显示相应信息
            if (_wasCancelled)
            {
                AppendLog($"⚠️ 转换被用户取消", LogLevel.Warning);
            }
            else if (failedCount > 0)
            {
                AppendLog($"❌ 部分文件转换失败", LogLevel.Error);
            }
            else if (_totalFilesCount > 0)
            {
                AppendLog($"🎉 恭喜！所有文件转换成功！");
            }

            if (_totalFilesCount == 0)
            {
                AppendLog($"⚠ 未找到需要转换的文件", LogLevel.Warning);
            }

            AppendLog($"==========================================");
        }

        private async Task DeleteOriginalFilesAsync()
        {
            if (_successfullyConvertedFiles.Count == 0)
                return;

            AppendLog($"==============开始删除原文件==============");
            AppendLog($"准备删除文件，释放资源中...");

            // 适度等待COM对象释放，减少不必要的延迟
            await Task.Delay(2000);

            // 执行一次垃圾回收确保COM对象释放
            GC.Collect();
            GC.WaitForPendingFinalizers();
            await Task.Delay(500);

            AppendLog($"开始删除文件...");

            var filesToDelete = new List<string>(_successfullyConvertedFiles);
            _successfullyConvertedFiles.Clear();

            int deletedCount = 0;
            var failedFiles = new List<string>();

            foreach (var filePath in filesToDelete)
            {
                if (!File.Exists(filePath))
                {
                    AppendLog($"✓ 文件已不存在: {Path.GetFileName(filePath)}");
                    continue;
                }

                bool deleted = false;

                // 尝试删除文件，逐步增加等待时间
                for (int attempt = 0; attempt < 5 && !deleted; attempt++)
                {
                    try
                    {
                        if (attempt > 0)
                        {
                            // 渐进式等待：1秒、2秒、3秒、4秒
                            int waitTime = attempt * 1000;
                            AppendLog($"⏳ 第{attempt + 1}次尝试删除: {Path.GetFileName(filePath)} (等待{attempt}秒)");
                            await Task.Delay(waitTime);

                            // 重试前再次垃圾回收
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }

                        // 先尝试修改文件属性，移除只读等限制
                        File.SetAttributes(filePath, FileAttributes.Normal);

                        // 在删除前备份文件，以便撤回时恢复
                        string backupPath = null;
                        try
                        {
                            backupPath = CreateBackupFile(filePath);
                        }
                        catch (Exception backupEx)
                        {
                            AppendLog($"⚠ 备份文件失败: {Path.GetFileName(filePath)} - {backupEx.Message}", LogLevel.Warning);
                            // 备份失败则跳过删除，避免不可恢复的数据丢失
                            continue;
                        }

                        // 删除文件
                        File.Delete(filePath);
                        AppendLog($"✓ 原文件已删除: {Path.GetFileName(filePath)}");

                        // 记录删除操作到撤回历史
                        RecordDeleteOperation(filePath, backupPath);

                        deletedCount++;
                        deleted = true;
                    }
                    catch (IOException) when (attempt < 4)
                    {
                        if (attempt == 3) // 最后一次尝试前，强制清理进程
                        {
                            AppendLog($"⚠ 常规方式删除失败，尝试清理相关进程...");
                            await ForceCleanupOfficeProcesses();
                        }
                    }
                    catch (UnauthorizedAccessException) when (attempt < 4)
                    {
                        // 权限问题也重试
                    }
                    catch (Exception ex)
                    {
                        if (attempt == 4)
                        {
                            AppendLog($"✗ 删除失败: {Path.GetFileName(filePath)} - {ex.Message}", LogLevel.Warning);
                            failedFiles.Add(filePath);
                        }
                    }
                }
            }

            if (failedFiles.Count > 0)
            {
                AppendLog($"⚠ 成功删除 {deletedCount} 个文件，{failedFiles.Count} 个文件删除失败:", LogLevel.Warning);
                foreach (var filePath in failedFiles)
                {
                    AppendLog($"   - {Path.GetFileName(filePath)}", LogLevel.Warning);
                }
                AppendLog($"💡 提示：请手动删除这些文件，或检查文件权限设置。", LogLevel.Info);
            }
            else
            {
                AppendLog($"✅ 成功删除所有 {deletedCount} 个原文件!");
            }

            AppendLog($"==============文件删除完成==============");
        }

        private async Task ForceCleanupOfficeProcesses()
        {
            try
            {
                AppendLog($"正在清理残留的Office进程...");

                // 强制终止Office和WPS进程（只在删除失败时作为最后手段）
                var processNames = new[] {
                    "WINWORD", "EXCEL", "POWERPNT",  // MS Office
                    "wps", "et", "wpp"               // WPS Office
                };

                foreach (var processName in processNames)
                {
                    var processes = Process.GetProcessesByName(processName);
                    foreach (var process in processes)
                    {
                        try
                        {
                            if (!process.HasExited)
                            {
                                process.Kill();
                                AppendLog($"强制终止残留进程: {processName} (PID: {process.Id})");
                                await Task.Delay(500);
                            }
                        }
                        catch
                        {
                            // 忽略进程终止错误
                        }
                    }
                }

                // 等待进程完全退出
                await Task.Delay(2000);
            }
            catch (Exception ex)
            {
                AppendLog($"清理Office进程时出错: {ex.Message}", LogLevel.Warning);
            }
        }

        private bool IsWpsOfficeAvailable()
        {
            try
            {
                // 尝试创建WPS应用程序实例来检测可用性（使用正确的 ProgID 集合）
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

        // 撤回功能相关方法
        private void RecordConversionOperation(string targetFilePath, string sourceFilePath, DuplicateFileAction action)
        {
            var operation = new ConversionOperation
            {
                FilePath = targetFilePath,
                SourceFile = sourceFilePath,
                Timestamp = DateTime.Now
            };

            // 根据操作类型记录不同的信息
            switch (action)
            {
                case DuplicateFileAction.Overwrite:
                    operation.Type = OperationType.OverwriteFile;
                    // 备份原文件到临时位置
                    if (File.Exists(targetFilePath))
                    {
                        operation.BackupPath = CreateBackupFile(targetFilePath);
                        // 如果备份失败，发出警告
                        if (string.IsNullOrEmpty(operation.BackupPath))
                        {
                            AppendLog($"警告: 文件 {Path.GetFileName(targetFilePath)} 备份失败，撤回时无法恢复原文件", LogLevel.Warning);
                        }
                    }
                    break;
                case DuplicateFileAction.Rename:
                case DuplicateFileAction.Skip:
                default:
                    operation.Type = OperationType.CreateFile;
                    break;
            }

            _conversionHistory.Add(operation);
        }

        private void RecordDeleteOperation(string deletedFilePath, string backupPath)
        {
            var operation = new ConversionOperation
            {
                Type = OperationType.DeleteFile,
                FilePath = deletedFilePath,
                BackupPath = backupPath,
                Timestamp = DateTime.Now
            };
            _conversionHistory.Add(operation);
        }

        private void RecordDirectoryCreation(string directoryPath)
        {
            var operation = new ConversionOperation
            {
                Type = OperationType.CreateDirectory,
                FilePath = directoryPath,
                Timestamp = DateTime.Now
            };
            _conversionHistory.Add(operation);
        }

        private string CreateBackupFile(string originalFilePath)
        {
            try
            {
                // 使用会话ID确保只清理当前会话的备份文件
                var tempDir = Path.Combine(Path.GetTempPath(), "Office2PDF_Backup", _sessionId);

                if (!Directory.Exists(tempDir))
                {
                    Directory.CreateDirectory(tempDir);
                }

                var backupPath = Path.Combine(tempDir, Path.GetFileName(originalFilePath));

                // 如果文件名仍然冲突，添加时间戳标识
                int counter = 1;
                var originalBackupPath = backupPath;
                while (File.Exists(backupPath))
                {
                    var fileName = Path.GetFileNameWithoutExtension(originalFilePath);
                    var extension = Path.GetExtension(originalFilePath);
                    var timestamp = DateTime.Now.ToString("HHmmss_fff");
                    backupPath = Path.Combine(tempDir, $"{fileName}_{timestamp}_{counter}{extension}");
                    counter++;
                }

                File.Copy(originalFilePath, backupPath, true);
                return backupPath;
            }
            catch (Exception ex)
            {
                AppendLog($"创建备份文件失败: {ex.Message}", LogLevel.Warning);
                return null;
            }
        }

        private async void UndoChanges_Click(object sender, RoutedEventArgs e)
        {
            if (_conversionHistory.Count == 0)
            {
                AppendLog("没有可撤回的操作", LogLevel.Info);
                return;
            }

            var result = System.Windows.MessageBox.Show(
                $"确定要撤回本次转换的所有更改吗？\n\n将会删除 {_conversionHistory.Count} 个转换生成的文件。\n此操作不可逆！",
                "确认撤回",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                await PerformUndo();
            }
        }

        private async Task PerformUndo()
        {
            UndoButton.IsEnabled = false;
            var undoneCount = 0;
            var failedCount = 0;

            AppendLog("开始撤回操作...", LogLevel.Info);

            // 统计各种操作类型
            var createCount = _conversionHistory.Count(op => op.Type == OperationType.CreateFile);
            var overwriteCount = _conversionHistory.Count(op => op.Type == OperationType.OverwriteFile);
            var deleteCount = _conversionHistory.Count(op => op.Type == OperationType.DeleteFile);
            var dirCount = _conversionHistory.Count(op => op.Type == OperationType.CreateDirectory);

            AppendLog($"将撤回：创建文件 {createCount} 个，覆盖文件 {overwriteCount} 个，删除原文件 {deleteCount} 个，创建目录 {dirCount} 个", LogLevel.Info);

            // 预检查：验证备份文件的完整性
            var brokenBackups = 0;
            foreach (var op in _conversionHistory)
            {
                if ((op.Type == OperationType.OverwriteFile || op.Type == OperationType.DeleteFile) &&
                    !string.IsNullOrEmpty(op.BackupPath) &&
                    !File.Exists(op.BackupPath))
                {
                    brokenBackups++;
                }
            }

            if (brokenBackups > 0)
            {
                AppendLog($"警告: 检测到 {brokenBackups} 个备份文件丢失，对应的覆盖/删除操作无法完全撤回", LogLevel.Warning);
            }

            // 按时间倒序撤回操作
            for (int i = _conversionHistory.Count - 1; i >= 0; i--)
            {
                var operation = _conversionHistory[i];
                try
                {
                    switch (operation.Type)
                    {
                        case OperationType.CreateFile:
                            // 删除创建的文件
                            if (File.Exists(operation.FilePath))
                            {
                                File.Delete(operation.FilePath);
                                AppendLog($"已删除: {Path.GetFileName(operation.FilePath)}");
                                undoneCount++;
                            }
                            break;

                        case OperationType.OverwriteFile:
                            // 删除新文件，恢复备份
                            if (File.Exists(operation.FilePath))
                            {
                                File.Delete(operation.FilePath);
                            }
                            if (!string.IsNullOrEmpty(operation.BackupPath) && File.Exists(operation.BackupPath))
                            {
                                File.Copy(operation.BackupPath, operation.FilePath, true);
                                AppendLog($"已恢复: {Path.GetFileName(operation.FilePath)}");
                                undoneCount++;
                            }
                            else if (string.IsNullOrEmpty(operation.BackupPath))
                            {
                                AppendLog($"无法恢复 {Path.GetFileName(operation.FilePath)}: 备份文件不存在", LogLevel.Warning);
                                failedCount++;
                            }
                            else
                            {
                                AppendLog($"无法恢复 {Path.GetFileName(operation.FilePath)}: 备份文件已损坏或被删除", LogLevel.Error);
                                failedCount++;
                            }
                            break;

                        case OperationType.CreateDirectory:
                            // 删除创建的目录（仅当目录为空时）
                            if (Directory.Exists(operation.FilePath))
                            {
                                try
                                {
                                    // 只删除空目录，避免误删有其他文件的目录
                                    if (Directory.GetFiles(operation.FilePath).Length == 0 &&
                                        Directory.GetDirectories(operation.FilePath).Length == 0)
                                    {
                                        Directory.Delete(operation.FilePath);
                                        AppendLog($"已删除空目录: {Path.GetFileName(operation.FilePath)}");
                                        undoneCount++;
                                    }
                                    else
                                    {
                                        AppendLog($"目录非空，跳过删除: {Path.GetFileName(operation.FilePath)}", LogLevel.Info);
                                    }
                                }
                                catch (Exception dirEx)
                                {
                                    AppendLog($"删除目录失败: {Path.GetFileName(operation.FilePath)} - {dirEx.Message}", LogLevel.Warning);
                                }
                            }
                            break;

                        case OperationType.DeleteFile:
                            // 恢复被删除的原文件
                            if (!string.IsNullOrEmpty(operation.BackupPath) && File.Exists(operation.BackupPath))
                            {
                                try
                                {
                                    // 确保目标目录存在
                                    string targetDir = Path.GetDirectoryName(operation.FilePath);
                                    if (!Directory.Exists(targetDir))
                                    {
                                        Directory.CreateDirectory(targetDir);
                                    }

                                    // 恢复被删除的文件
                                    File.Copy(operation.BackupPath, operation.FilePath, true);
                                    AppendLog($"已恢复被删除的文件: {Path.GetFileName(operation.FilePath)}");
                                    undoneCount++;
                                }
                                catch (Exception restoreEx)
                                {
                                    AppendLog($"恢复被删除文件失败: {Path.GetFileName(operation.FilePath)} - {restoreEx.Message}", LogLevel.Error);
                                    failedCount++;
                                }
                            }
                            else
                            {
                                AppendLog($"无法恢复被删除的文件 {Path.GetFileName(operation.FilePath)}: 备份文件不存在", LogLevel.Error);
                                failedCount++;
                            }
                            break;
                    }
                }
                catch (Exception ex)
                {
                    AppendLog($"撤回失败 {Path.GetFileName(operation.FilePath)}: {ex.Message}", LogLevel.Error);
                    failedCount++;
                }

                // 避免UI冻结
                if (i % 10 == 0)
                {
                    await Task.Delay(1);
                }
            }

            // 清理备份文件
            CleanupBackupFiles();

            // 清除历史记录
            _conversionHistory.Clear();

            AppendLog($"撤回完成: 成功 {undoneCount} 个，失败 {failedCount} 个",
                failedCount > 0 ? LogLevel.Warning : LogLevel.Info);
        }

        private void CleanupCurrentSessionBackups()
        {
            try
            {
                // 只清理当前会话的备份文件
                foreach (var operation in _conversionHistory)
                {
                    if (!string.IsNullOrEmpty(operation.BackupPath) && File.Exists(operation.BackupPath))
                    {
                        File.Delete(operation.BackupPath);
                    }
                }

                // 清理当前会话的临时目录
                var sessionTempDir = Path.Combine(Path.GetTempPath(), "Office2PDF_Backup", _sessionId);
                if (Directory.Exists(sessionTempDir))
                {
                    try
                    {
                        Directory.Delete(sessionTempDir, true);
                    }
                    catch
                    {
                        // 忽略清理失败
                    }
                }
            }
            catch
            {
                // 清理失败不影响主要功能
            }
        }

        private void CleanupBackupFiles()
        {
            try
            {
                // 清理当前会话的备份文件
                CleanupCurrentSessionBackups();
            }
            catch
            {
                // 清理失败不影响主要功能
            }
        }

    }

    public class MainWindowViewModel : INotifyPropertyChanged
    {
        public MainWindowViewModel()
        {
            // 初始化时根据各个转换类型的状态更新"全选"状态
            UpdateIsConvertAll();
        }

        private string ProcessPath(string path)
        {
            if (!string.IsNullOrEmpty(path))
            {
                path = path.Trim();
                if (path.StartsWith("\"") && path.EndsWith("\""))
                {
                    path = path.Substring(1, path.Length - 2);
                }
            }
            return path;
        }

        private string _fromFolderPath = "";
        public string FromRootFolderPath
        {
            get => _fromFolderPath;
            set
            {
                var processedValue = ProcessPath(value);
                if (_fromFolderPath != processedValue)
                {
                    _fromFolderPath = processedValue;
                    OnPropertyChanged();

                    // 如果来源路径有效且目标路径为空，自动生成目标路径
                    if (!string.IsNullOrWhiteSpace(processedValue) &&
                        Directory.Exists(processedValue) &&
                        string.IsNullOrWhiteSpace(_toFolderPath))
                    {
                        ToRootFolderPath = processedValue + "_PDFs";
                    }

                    // 通知主窗口更新按钮状态
                    System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        if (System.Windows.Application.Current.MainWindow is MainWindow mainWindow)
                        {
                            mainWindow.UpdateCanStartConvert();
                        }
                    }));
                }
            }
        }

        private string _toFolderPath = "";
        public string ToRootFolderPath
        {
            get => _toFolderPath;
            set
            {
                var processedValue = ProcessPath(value);
                if (_toFolderPath != processedValue)
                {
                    _toFolderPath = processedValue;
                    OnPropertyChanged();
                    // 通知主窗口更新按钮状态
                    System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        if (System.Windows.Application.Current.MainWindow is MainWindow mainWindow)
                        {
                            mainWindow.UpdateCanStartConvert();
                        }
                    }));
                }
            }
        }

        private bool _isConvertWord = true;
        public bool IsConvertWord
        {
            get => _isConvertWord;
            set
            {
                if (_isConvertWord != value)
                {
                    _isConvertWord = value;
                    OnPropertyChanged();
                    UpdateIsConvertAll();
                }
            }
        }

        private bool _isConvertPPT = true;
        public bool IsConvertPPT
        {
            get => _isConvertPPT;
            set
            {
                if (_isConvertPPT != value)
                {
                    _isConvertPPT = value;
                    OnPropertyChanged();
                    UpdateIsConvertAll();
                }
            }
        }

        private bool _isConvertExcel = false;  // 默认不勾选Excel
        public bool IsConvertExcel
        {
            get => _isConvertExcel;
            set
            {
                if (_isConvertExcel != value)
                {
                    _isConvertExcel = value;
                    OnPropertyChanged();
                    UpdateIsConvertAll();
                }
            }
        }

        private bool _isConvertAll = false;
        public bool IsConvertAll
        {
            get => _isConvertAll;
            set
            {
                if (_isConvertAll != value)
                {
                    _isConvertAll = value;
                    _isConvertWord = value;
                    _isConvertPPT = value;
                    _isConvertExcel = value;

                    OnPropertyChanged();
                    OnPropertyChanged(nameof(IsConvertWord));
                    OnPropertyChanged(nameof(IsConvertPPT));
                    OnPropertyChanged(nameof(IsConvertExcel));
                }
            }
        }

        private bool _isConvertChildrenFolder = true;
        public bool IsConvertChildrenFolder
        {
            get => _isConvertChildrenFolder;
            set
            {
                if (_isConvertChildrenFolder != value)
                {
                    _isConvertChildrenFolder = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _isKeepFolderStructure = true;
        public bool IsKeepFolderStructure
        {
            get => _isKeepFolderStructure;
            set
            {
                if (_isKeepFolderStructure != value)
                {
                    _isKeepFolderStructure = value;
                    OnPropertyChanged();
                }
            }
        }


        private bool _isPrintRevisionsInWord = false;
        public bool IsPrintRevisionsInWord
        {
            get => _isPrintRevisionsInWord;
            set
            {
                if (_isPrintRevisionsInWord != value)
                {
                    _isPrintRevisionsInWord = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _isConvertOneSheetOnePDFInExcel = false;
        public bool IsConvertOneSheetOnePDFInExcel
        {
            get => _isConvertOneSheetOnePDFInExcel;
            set
            {
                if (_isConvertOneSheetOnePDFInExcel != value)
                {
                    _isConvertOneSheetOnePDFInExcel = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _isDeleteOriginalFiles = false;
        public bool IsDeleteOriginalFiles
        {
            get => _isDeleteOriginalFiles;
            set
            {
                // 如果用户尝试勾选删除原文件选项，弹出确认对话框
                if (!_isDeleteOriginalFiles && value)
                {
                    var result = System.Windows.MessageBox.Show(
                        "警告：此操作将在转换完成后删除所有成功转换的原文件！\n\n" +
                        "被删除的文件将自动备份，可通过\"撤回更改\"功能恢复。\n" +
                        "请确保有足够的磁盘空间用于备份文件。\n\n" +
                        "您确定要启用删除原文件功能吗？",
                        "删除原文件确认",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Warning,
                        MessageBoxResult.No);

                    if (result != MessageBoxResult.Yes)
                    {
                        // 用户选择了取消，不改变状态
                        OnPropertyChanged(); // 通知UI更新，保持复选框为未选中状态
                        return;
                    }
                }

                if (_isDeleteOriginalFiles != value)
                {
                    _isDeleteOriginalFiles = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _useWpsOffice = true;  // 默认选择WPS Office
        public bool UseWpsOffice
        {
            get => _useWpsOffice;
            set
            {
                if (_useWpsOffice != value)
                {
                    _useWpsOffice = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _canStartConvert = true;
        public bool CanStartConvert
        {
            get => _canStartConvert;
            set
            {
                if (_canStartConvert != value)
                {
                    _canStartConvert = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _buttonText = "开始";
        public string ButtonText
        {
            get => _buttonText;
            set
            {
                if (_buttonText != value)
                {
                    _buttonText = value;
                    OnPropertyChanged();
                }
            }
        }

        private void UpdateIsConvertAll()
        {
            _isConvertAll = IsConvertWord && IsConvertPPT && IsConvertExcel;
            OnPropertyChanged(nameof(IsConvertAll));
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}

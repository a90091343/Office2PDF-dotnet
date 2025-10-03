using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace Office2PDF
{
    /// <summary>
    /// 转换引擎：负责所有PDF转换相关的业务逻辑
    /// </summary>
    public class ConversionEngine
    {
        // 常量定义
        private const int PPT_PROCESS_STABILIZATION_DELAY_MS = 500; // PPT进程需要的稳定延迟
        private const int MAX_FILENAME_RETRY_ATTEMPTS = 9999; // 文件名重命名最大尝试次数
        private const int WINDOWS_MAX_PATH_LENGTH = 260; // Windows路径最大长度限制
        private const int MIN_FILENAME_LENGTH = 20; // 截断文件名时保留的最小长度

        // 状态跟踪（线程安全集合）
        private ConcurrentDictionary<string, byte> _successfullyConvertedFiles = new ConcurrentDictionary<string, byte>();
        private ConcurrentDictionary<string, byte> _processedSuccessfulFiles = new ConcurrentDictionary<string, byte>(); // 记录已成功处理的文件，避免重复统计
        private ConcurrentDictionary<string, byte> _failedFiles = new ConcurrentDictionary<string, byte>();
        private ConcurrentDictionary<string, byte> _skippedFiles = new ConcurrentDictionary<string, byte>();
        private ConcurrentDictionary<string, byte> _overwrittenFiles = new ConcurrentDictionary<string, byte>();
        private ConcurrentDictionary<string, byte> _renamedFiles = new ConcurrentDictionary<string, byte>();
        private ConcurrentDictionary<Type, byte> _engineInfoLoggedFor = new ConcurrentDictionary<Type, byte>();
        private int _batchModeProcessedCount = 0;
        private ConcurrentDictionary<string, byte> _conflictFiles = new ConcurrentDictionary<string, byte>();

        // 统计数据（使用 Interlocked 进行原子操作）
        private int _successfulWordCount = 0;
        private int _successfulExcelCount = 0;
        private int _successfulPptCount = 0;
        private int _totalWordCount = 0;
        private int _totalExcelCount = 0;
        private int _totalPptCount = 0;
        private int _totalFilesCount = 0;
        private bool _wasCancelled = false;

        // 配置和设置
        private DuplicateFileAction _duplicateFileAction = DuplicateFileAction.Skip;
        private readonly MainWindowViewModel _viewModel;
        private readonly Action<string, LogLevel> _logAction;
        private CancellationToken _cancellationToken;

        // 撤回功能（线程安全集合）
        private ConcurrentBag<ConversionOperation> _conversionHistory = new ConcurrentBag<ConversionOperation>();
        private readonly string _sessionId = DateTime.Now.ToString("yyyyMMdd_HHmmss_") + Guid.NewGuid().ToString("N").Substring(0, 8);

        public ConversionEngine(MainWindowViewModel viewModel, Action<string, LogLevel> logAction)
        {
            _viewModel = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            _logAction = logAction ?? throw new ArgumentNullException(nameof(logAction));
        }

        /// <summary>
        /// 获取转换历史记录数量
        /// </summary>
        public int ConversionHistoryCount => _conversionHistory.Count;

        /// <summary>
        /// 获取成功转换的文件列表（用于删除原文件）
        /// </summary>
        public IReadOnlyCollection<string> SuccessfullyConvertedFiles => _successfullyConvertedFiles.Keys.ToList();

        /// <summary>
        /// 重置引擎状态，准备新的转换任务
        /// </summary>
        public void Reset()
        {
            _successfullyConvertedFiles.Clear();
            _processedSuccessfulFiles.Clear();
            _failedFiles.Clear();
            _skippedFiles.Clear();
            _overwrittenFiles.Clear();
            _renamedFiles.Clear();
            _totalFilesCount = 0;
            _wasCancelled = false;
            Interlocked.Exchange(ref _successfulWordCount, 0);
            Interlocked.Exchange(ref _successfulExcelCount, 0);
            Interlocked.Exchange(ref _successfulPptCount, 0);
            Interlocked.Exchange(ref _totalWordCount, 0);
            Interlocked.Exchange(ref _totalExcelCount, 0);
            Interlocked.Exchange(ref _totalPptCount, 0);
            _batchModeProcessedCount = 0;
            _engineInfoLoggedFor.Clear();
            _conflictFiles.Clear();
        }

        /// <summary>
        /// 设置重复文件处理策略
        /// </summary>
        public void SetDuplicateFileAction(DuplicateFileAction action)
        {
            _duplicateFileAction = action;
        }

        /// <summary>
        /// 清除撤回历史
        /// </summary>
        public void ClearConversionHistory()
        {
            CleanupBackupFiles();
            // ConcurrentBag 没有 Clear 方法，重新创建实例
            _conversionHistory = new ConcurrentBag<ConversionOperation>();
        }

        /// <summary>
        /// 预扫描所有待转换文件，检测文件名冲突
        /// </summary>
        public void PreScanFilesForConflicts<TConvertAction>((string TypeName, bool IsConvert, string[] Extensions, TConvertAction ConvertAction)[] fileTypeHandlers)
        {
            // 按目录分组，存储每个目录下的文件基本名（不含扩展名）
            var filesByDirectory = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            var searchOption = _viewModel.IsConvertChildrenFolder ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

            // 第一遍：收集所有待转换文件的信息
            var allFilesInfo = new List<(string fullPath, string directory, string baseName, string extension)>();

            foreach (var (TypeName, IsConvert, Extensions, ConvertAction) in fileTypeHandlers)
            {
                if (!IsConvert) continue;

                var files = Directory.EnumerateFiles(_viewModel.FromRootFolderPath, "*.*", searchOption)
                    .Where(file => Extensions.Any(ext => file.EndsWith(ext, StringComparison.OrdinalIgnoreCase)))
                    .ToArray();

                foreach (var file in files)
                {
                    var directory = Path.GetDirectoryName(file);
                    var baseName = Path.GetFileNameWithoutExtension(file);
                    var extension = Path.GetExtension(file);

                    allFilesInfo.Add((file, directory, baseName, extension));
                }
            }

            // 第二遍：检测每个目录下是否有同名的不同扩展名文件
            foreach (var fileInfo in allFilesInfo)
            {
                var targetDir = fileInfo.directory;
                if (_viewModel.IsKeepFolderStructure)
                {
                    // 如果保持目录结构，需要计算对应的目标目录
                    var relativePath = GetRelativePath(_viewModel.FromRootFolderPath, fileInfo.directory);
                    targetDir = Path.Combine(_viewModel.ToRootFolderPath, relativePath);
                }
                else
                {
                    targetDir = _viewModel.ToRootFolderPath;
                }

                // 为该目标目录创建文件基本名集合
                if (!filesByDirectory.ContainsKey(targetDir))
                {
                    filesByDirectory[targetDir] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                }

                var baseName = fileInfo.baseName;

                // 如果该目录下已经有相同基本名的文件，说明有冲突
                if (filesByDirectory[targetDir].Contains(baseName))
                {
                    // 标记当前文件需要保留完整扩展名
                    _conflictFiles.TryAdd(fileInfo.fullPath, 0);

                    // 同时也需要找到之前的同名文件并标记它
                    var previousFile = allFilesInfo.FirstOrDefault(f =>
                    {
                        var prevTargetDir = f.directory;
                        if (_viewModel.IsKeepFolderStructure)
                        {
                            var relativePath = GetRelativePath(_viewModel.FromRootFolderPath, f.directory);
                            prevTargetDir = Path.Combine(_viewModel.ToRootFolderPath, relativePath);
                        }
                        else
                        {
                            prevTargetDir = _viewModel.ToRootFolderPath;
                        }

                        return prevTargetDir.Equals(targetDir, StringComparison.OrdinalIgnoreCase) &&
                               f.baseName.Equals(baseName, StringComparison.OrdinalIgnoreCase) &&
                               f.fullPath != fileInfo.fullPath;
                    });

                    if (previousFile.fullPath != null)
                    {
                        _conflictFiles.TryAdd(previousFile.fullPath, 0);
                    }
                }
                else
                {
                    filesByDirectory[targetDir].Add(baseName);
                }
            }

            // 如果检测到冲突，输出日志
            if (_conflictFiles.Count > 0)
            {
                Log($"⚠️ 检测到 {_conflictFiles.Count} 个文件名冲突，将保留完整扩展名（如：1.doc.pdf、1.ppt.pdf）", LogLevel.Warning);
            }
        }

        /// <summary>
        /// 转换为PDF（主入口）
        /// </summary>
        public async Task ConvertToPDFAsync<T>(string typeName, string[] fromFilePaths, CancellationToken cancellationToken) where T : IOfficeApplication, new()
        {
            _cancellationToken = cancellationToken;
            _batchModeProcessedCount = 0;

            // 尝试批量模式（使用单个进程处理所有文件）
            try
            {
                ConvertToPDFBatchMode<T>(typeName, fromFilePaths);
            }
            catch (OperationCanceledException)
            {
                // 用户取消，不切换到安全模式，直接退出
                return;
            }
            catch (Exception)
            {
                // 只用一句话提示用户
                Log($"🔄 检测到进程异常，自动切换到逐个文件处理模式...", LogLevel.Info);

                // 计算剩余未处理的文件
                if (_batchModeProcessedCount < fromFilePaths.Length)
                {
                    var remainingFiles = fromFilePaths.Skip(_batchModeProcessedCount).ToArray();

                    // 只处理剩余的文件
                    if (remainingFiles.Length > 0)
                    {
                        await ConvertToPDFSafeModeAsync<T>(typeName, remainingFiles, _batchModeProcessedCount);
                    }
                }
            }
        }

        /// <summary>
        /// 自动模式转换PDF，支持MS Office到WPS的自动回退
        /// </summary>
        public async Task ConvertToPDFWithAutoFallbackAsync<TPrimary, TFallback>(string typeName, string[] fromFilePaths, CancellationToken cancellationToken)
            where TPrimary : IOfficeApplication, new()
            where TFallback : IOfficeApplication, new()
        {
            _cancellationToken = cancellationToken;

            try
            {
                // 首先尝试使用MS Office引擎
                await ConvertToPDFAsync<TPrimary>(typeName, fromFilePaths, cancellationToken);
            }
            catch (OperationCanceledException)
            {
                // 用户取消，不要尝试回退，直接退出
                throw;
            }
            catch (System.Runtime.InteropServices.COMException ex) when (ex.HResult == unchecked((int)0x80040154))
            {
                // COM组件未注册错误，自动切换到WPS引擎
                Log($"❌ MS Office {typeName} COM组件未注册", LogLevel.Warning);
                Log($"🔄 自动切换到 WPS Office {typeName} 引擎进行重试...", LogLevel.Info);

                try
                {
                    await ConvertToPDFAsync<TFallback>(typeName, fromFilePaths, cancellationToken);
                    Log($"✅ 使用 WPS Office {typeName} 引擎转换成功", LogLevel.Info);
                }
                catch (Exception fallbackEx)
                {
                    Log($"❌ WPS Office {typeName} 引擎也转换失败: {fallbackEx.Message}", LogLevel.Error);
                    throw;
                }
            }
            catch (Exception ex) when (ex.Message.Contains("80040154") || ex.Message.Contains("没有注册类"))
            {
                // 其他形式的COM组件未注册错误
                Log($"❌ MS Office {typeName} COM组件未注册", LogLevel.Warning);
                Log($"🔄 自动切换到 WPS Office {typeName} 引擎进行重试...", LogLevel.Info);

                try
                {
                    await ConvertToPDFAsync<TFallback>(typeName, fromFilePaths, cancellationToken);
                    Log($"✅ 使用 WPS Office {typeName} 引擎转换成功", LogLevel.Info);
                }
                catch (Exception fallbackEx)
                {
                    Log($"❌ WPS Office {typeName} 引擎也转换失败: {fallbackEx.Message}", LogLevel.Error);
                    throw;
                }
            }
        }

        /// <summary>
        /// 批量模式：使用单个进程处理所有文件（高性能）
        /// </summary>
        private void ConvertToPDFBatchMode<T>(string typeName, string[] fromFilePaths) where T : IOfficeApplication, new()
        {
            var numberFormat = $"D{fromFilePaths.Length.ToString().Length}";

            // 创建单个应用程序实例
            IOfficeApplication application = CreateApplication<T>();
            ConfigureApplicationProperties(application);

            try
            {
                using (application)
                {
                    for (int i = 0; i < fromFilePaths.Length; i++)
                    {
                        if (_cancellationToken.IsCancellationRequested)
                        {
                            _wasCancelled = true;
                            Log($"{typeName} 转换已被用户取消", LogLevel.Warning);
                            break;
                        }

                        var index = i + 1;
                        var fromFilePath = fromFilePaths[i];

                        try
                        {
                            ProcessSingleFile(application, typeName, fromFilePath, index, numberFormat);
                            _batchModeProcessedCount++;
                        }
                        catch (OperationCanceledException)
                        {
                            // 用户取消操作，不算失败，直接退出
                            _wasCancelled = true;
                            break;
                        }
                        catch (Exception ex)
                        {
                            if (IsProcessCriticalError(ex))
                            {
                                throw;
                            }
                            else
                            {
                                _failedFiles.TryAdd(fromFilePath, 0);
                                Log($"（{index.ToString(numberFormat)}）{typeName} 转换出错：{fromFilePath} {ex.Message}", LogLevel.Error);
                                _batchModeProcessedCount++;
                            }
                        }
                    }

                    if (_wasCancelled)
                        Log($"{typeName} 类型文件转换被中途取消");
                }
            }
            finally
            {
                // 确保资源释放
            }
        }

        /// <summary>
        /// 安全模式：为每个文件创建和销毁进程（稳定但较慢）
        /// </summary>
        private async Task ConvertToPDFSafeModeAsync<T>(string typeName, string[] fromFilePaths, int startIndex = 0) where T : IOfficeApplication, new()
        {
            var totalCount = startIndex + fromFilePaths.Length;
            var numberFormat = $"D{totalCount.ToString().Length}";

            for (int i = 0; i < fromFilePaths.Length; i++)
            {
                if (_cancellationToken.IsCancellationRequested)
                {
                    _wasCancelled = true;
                    Log($"{typeName} 转换已被用户取消", LogLevel.Warning);
                    break;
                }

                var index = startIndex + i + 1;
                var fromFilePath = fromFilePaths[i];
                bool convertSuccess = false;
                int retryCount = 0;
                const int maxRetries = 3;

                // 重试机制：最多尝试3次
                while (!convertSuccess && retryCount < maxRetries)
                {
                    // 在重试循环开始时检查取消
                    if (_cancellationToken.IsCancellationRequested)
                    {
                        _wasCancelled = true;
                        break;
                    }

                    IOfficeApplication application = null;

                    try
                    {
                        // 创建和配置应用程序（可能失败）
                        application = CreateApplication<T>();
                        ConfigureApplicationProperties(application);

                        using (application)
                        {
                            ProcessSingleFile(application, typeName, fromFilePath, index, numberFormat);

                            // 专门为不稳定的PPT/WPP进程添加延迟
                            if (application is MSPowerPointApplication || application is WpsPresentationApplication)
                            {
                                await Task.Delay(PPT_PROCESS_STABILIZATION_DELAY_MS);
                            }
                        }

                        // 如果执行到这里说明转换成功或被跳过
                        convertSuccess = true;
                    }
                    catch (OperationCanceledException)
                    {
                        // 用户取消操作，这是正常行为，不算失败，不需要重试
                        _wasCancelled = true;

                        // 确保资源释放
                        if (application != null)
                        {
                            try
                            {
                                application.Dispose();
                            }
                            catch
                            {
                                // 忽略释放资源时的异常
                            }
                        }

                        // 跳出重试循环和文件循环
                        break;
                    }
                    catch (Exception ex)
                    {
                        retryCount++;

                        // 确保资源释放
                        if (application != null)
                        {
                            try
                            {
                                application.Dispose();
                            }
                            catch
                            {
                                // 忽略释放资源时的异常
                            }
                        }

                        // 检查文件是否已经成功处理（ProcessSingleFile已完成）
                        if (_processedSuccessfulFiles.ContainsKey(fromFilePath))
                        {
                            // 文件已成功转换，异常发生在cleanup阶段，忽略该异常
                            convertSuccess = true;
                            break;
                        }

                        // 如果还有重试机会
                        if (retryCount < maxRetries)
                        {
                            Log($"（{index.ToString(numberFormat)}）{typeName} 转换失败，正在重试 ({retryCount}/{maxRetries}): {ex.Message}", LogLevel.Warning);
                            await Task.Delay(1000); // 等待1秒后重试
                        }
                        else
                        {
                            // 重试次数用尽，记录为失败
                            if (!_failedFiles.ContainsKey(fromFilePath))
                            {
                                _failedFiles.TryAdd(fromFilePath, 0);
                            }
                            Log($"（{index.ToString(numberFormat)}）{typeName} 转换失败（已重试{maxRetries}次）: {fromFilePath} - {ex.Message}", LogLevel.Error);
                        }
                    }
                }

                // 如果用户取消了操作，跳出文件循环
                if (_wasCancelled)
                {
                    break;
                }
            }

            if (_wasCancelled)
                Log($"{typeName} 类型文件转换被中途取消");
        }

        /// <summary>
        /// 创建Office应用程序实例
        /// </summary>
        private IOfficeApplication CreateApplication<T>() where T : IOfficeApplication, new()
        {
            IOfficeApplication application;

            if (!_engineInfoLoggedFor.ContainsKey(typeof(T)))
            {
                // 第一次创建，使用带日志回调的构造函数
                Action<string> logCallback = (msg) =>
                {
                    if (msg.StartsWith("💡 提示:"))
                        Log(msg, LogLevel.Info);
                    else
                        Log(msg);
                };

                if (typeof(T) == typeof(MSWordApplication)) { application = new MSWordApplication(logCallback); }
                else if (typeof(T) == typeof(MSExcelApplication)) { application = new MSExcelApplication(logCallback); }
                else if (typeof(T) == typeof(MSPowerPointApplication)) { application = new MSPowerPointApplication(logCallback); }
                else { application = new T(); }

                _engineInfoLoggedFor.TryAdd(typeof(T), 0);
            }
            else
            {
                // 非第一次创建，使用不带日志回调的构造函数
                if (typeof(T) == typeof(MSWordApplication)) { application = new MSWordApplication(); }
                else if (typeof(T) == typeof(MSExcelApplication)) { application = new MSExcelApplication(); }
                else if (typeof(T) == typeof(MSPowerPointApplication)) { application = new MSPowerPointApplication(); }
                else { application = new T(); }
            }

            return application;
        }

        /// <summary>
        /// 配置应用程序属性（根据用户设置）
        /// </summary>
        private void ConfigureApplicationProperties(IOfficeApplication application)
        {
            if (application is MSWordApplication wordApp)
            {
                wordApp.IsPrintRevisions = _viewModel.IsPrintRevisionsInWord;
            }
            else if (application is WpsWriterApplication wpsWordApp)
            {
                wpsWordApp.IsPrintRevisions = _viewModel.IsPrintRevisionsInWord;
            }
            else if (application is MSExcelApplication excelApp)
            {
                excelApp.IsConvertOneSheetOnePDF = _viewModel.IsConvertOneSheetOnePDFInExcel;
            }
            else if (application is WpsSpreadsheetApplication wpsExcelApp)
            {
                wpsExcelApp.IsConvertOneSheetOnePDF = _viewModel.IsConvertOneSheetOnePDFInExcel;
            }
        }

        /// <summary>
        /// 处理单个文件的转换逻辑
        /// </summary>
        private void ProcessSingleFile(IOfficeApplication application, string typeName, string fromFilePath, int index, string numberFormat)
        {
            if (_cancellationToken.IsCancellationRequested)
            {
                _wasCancelled = true;
                Log($"{typeName} 转换已被用户取消", LogLevel.Warning);
                throw new OperationCanceledException();
            }

            application.OpenDocument(fromFilePath);

            try
            {
                var toFilePath = GetToFilePath(_viewModel.FromRootFolderPath, _viewModel.ToRootFolderPath, fromFilePath, Path.GetFileName(fromFilePath));

                // 处理重复文件
                var handleResult = HandleDuplicateFile(toFilePath);
                if (handleResult.FilePath == null)
                {
                    _skippedFiles.TryAdd(fromFilePath, 0);
                    Log($"（{index.ToString(numberFormat)}）{typeName} 已跳过: {Path.GetFileName(toFilePath)} (目标文件已存在)");
                    // 关闭已打开的文档
                    try
                    {
                        application.CloseDocument();
                    }
                    catch
                    {
                        // 忽略关闭文档时的异常
                    }
                    return;
                }

                if (_cancellationToken.IsCancellationRequested)
                {
                    _wasCancelled = true;
                    Log($"{typeName} 转换已被用户取消", LogLevel.Warning);
                    throw new OperationCanceledException();
                }

                // 记录处理类型
                if (!handleResult.IsOriginalFile)
                {
                    switch (handleResult.Action)
                    {
                        case DuplicateFileAction.Overwrite:
                            _overwrittenFiles.TryAdd(fromFilePath, 0);
                            break;
                        case DuplicateFileAction.Rename:
                            _renamedFiles.TryAdd(fromFilePath, 0);
                            break;
                    }
                }

                // 记录即将进行的操作用于撤回功能
                // 注意：对于 Excel Sheet 分离模式，推迟到 DetectExcelSheetFiles 后再记录
                // RecordConversionOperation(handleResult.FilePath, fromFilePath, handleResult.Action);

                // 执行转换
                application.SaveAsPDF(handleResult.FilePath);

                // 检查Excel Sheet分离模式
                List<string> actualGeneratedFiles = new List<string>();
                bool isExcelApplication = application is MSExcelApplication || application is WpsSpreadsheetApplication;

                if (isExcelApplication && _viewModel.IsConvertOneSheetOnePDFInExcel)
                {
                    actualGeneratedFiles = DetectExcelSheetFiles(handleResult.FilePath, fromFilePath, handleResult.Action);
                }
                else
                {
                    actualGeneratedFiles.Add(handleResult.FilePath);
                    // 非 Sheet 分离模式，记录单个文件操作
                    RecordConversionOperation(handleResult.FilePath, fromFilePath, handleResult.Action);
                }

                // 输出转换结果日志
                LogConversionResult(typeName, toFilePath, actualGeneratedFiles, index, numberFormat);

                // 统计成功转换的文件类型
                IncrementSuccessCount(fromFilePath);

                // 如果选择了删除原文件，则将文件路径添加到待删除列表
                if (_viewModel.IsDeleteOriginalFiles)
                {
                    _successfullyConvertedFiles.TryAdd(fromFilePath, 0);
                }
            }
            finally
            {
                try
                {
                    application.CloseDocument();
                }
                catch (Exception)
                {
                    // 忽略关闭文档时的异常
                }
            }
        }

        /// <summary>
        /// 检测Excel Sheet分离模式生成的文件，并处理重复文件
        /// </summary>
        private List<string> DetectExcelSheetFiles(string handleResultFilePath, string fromFilePath, DuplicateFileAction originalAction)
        {
            List<string> actualGeneratedFiles = new List<string>();
            var directory = Path.GetDirectoryName(handleResultFilePath);
            var baseFileName = Path.GetFileNameWithoutExtension(handleResultFilePath);

            if (Directory.Exists(directory))
            {
                // 检测所有以 baseFileName_ 开头的 PDF 文件（这些是 Sheet 文件）
                // MS Office 和 WPS 格式: filename_SheetName.pdf
                var sheetPattern = $"{baseFileName}_*.pdf";
                var sheetFiles = Directory.GetFiles(directory, sheetPattern);

                var allSheetFiles = sheetFiles
                    .Where(f => !f.Equals(handleResultFilePath, StringComparison.OrdinalIgnoreCase))
                    .Distinct()
                    .ToList();

                if (allSheetFiles.Count > 0)
                {
                    // Sheet 分离模式：只记录实际生成的 Sheet 文件，不记录原始文件

                    // 处理每个Sheet文件的重复问题
                    foreach (var sheetFile in allSheetFiles)
                    {
                        string finalSheetPath = sheetFile;
                        DuplicateFileAction actionTaken = originalAction;

                        // 检查是否需要处理重复文件（只在第二次及以后转换时）
                        if (_processedSuccessfulFiles.ContainsKey(fromFilePath))
                        {
                            var handleResult = HandleDuplicateFileForSheet(sheetFile);
                            if (handleResult.FilePath == null)
                            {
                                // 跳过这个Sheet文件
                                continue;
                            }

                            finalSheetPath = handleResult.FilePath;
                            actionTaken = handleResult.Action;

                            // 如果需要重命名或覆盖，则移动/重命名文件
                            if (!finalSheetPath.Equals(sheetFile, StringComparison.OrdinalIgnoreCase))
                            {
                                if (File.Exists(sheetFile))
                                {
                                    File.Move(sheetFile, finalSheetPath);
                                }
                            }

                            // 记录重命名或覆盖的文件
                            if (actionTaken == DuplicateFileAction.Rename)
                            {
                                _renamedFiles.TryAdd(fromFilePath, 0);
                            }
                            else if (actionTaken == DuplicateFileAction.Overwrite)
                            {
                                _overwrittenFiles.TryAdd(fromFilePath, 0);
                            }
                        }

                        actualGeneratedFiles.Add(finalSheetPath);
                        // 记录每个实际生成的 Sheet 文件
                        RecordConversionOperation(finalSheetPath, fromFilePath, actionTaken);
                    }
                }
                else
                {
                    // 没有检测到 Sheet 文件，使用原始文件
                    actualGeneratedFiles.Add(handleResultFilePath);
                    RecordConversionOperation(handleResultFilePath, fromFilePath, originalAction);
                }
            }
            else
            {
                actualGeneratedFiles.Add(handleResultFilePath);
                RecordConversionOperation(handleResultFilePath, fromFilePath, originalAction);
            }

            return actualGeneratedFiles;
        }

        /// <summary>
        /// 为Sheet文件处理重复文件（应用全局的重复文件策略）
        /// </summary>
        private FileHandleResult HandleDuplicateFileForSheet(string originalPath)
        {
            if (string.IsNullOrEmpty(originalPath))
            {
                throw new ArgumentException("文件路径不能为空", nameof(originalPath));
            }

            // 检查文件是否已存在于转换历史中（即之前已经生成过）
            bool isAlreadyConverted = _conversionHistory.Any(op =>
                op.FilePath != null && op.FilePath.Equals(originalPath, StringComparison.OrdinalIgnoreCase));

            if (!isAlreadyConverted)
            {
                // 这是第一次生成这个Sheet文件
                return new FileHandleResult
                {
                    FilePath = originalPath,
                    Action = DuplicateFileAction.Rename,
                    IsOriginalFile = true
                };
            }

            // 文件已经存在，应用用户选择的重复文件策略
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
                    return new FileHandleResult
                    {
                        FilePath = GetUniqueFilePath(originalPath),
                        Action = DuplicateFileAction.Rename,
                        IsOriginalFile = false
                    };
            }
        }

        /// <summary>
        /// 输出转换结果日志
        /// </summary>
        private void LogConversionResult(string typeName, string originalToFilePath, List<string> actualGeneratedFiles, int index, string numberFormat)
        {
            if (actualGeneratedFiles.Count == 1)
            {
                var generatedFile = actualGeneratedFiles[0];
                var logMessage = generatedFile == originalToFilePath
                    ? $"（{index.ToString(numberFormat)}）{typeName} 转换成功: {GetRelativePath(_viewModel.ToRootFolderPath, generatedFile)}"
                    : $"（{index.ToString(numberFormat)}）{typeName} 转换成功: {GetRelativePath(_viewModel.ToRootFolderPath, generatedFile)} (已重命名)";
                Log(logMessage);
            }
            else
            {
                Log($"（{index.ToString(numberFormat)}）{typeName} 转换成功，生成 {actualGeneratedFiles.Count} 个Sheet PDF:");
                foreach (var file in actualGeneratedFiles)
                {
                    Log($"    • {GetRelativePath(_viewModel.ToRootFolderPath, file)}");
                }
            }
        }

        /// <summary>
        /// 判断是否是需要切换到安全模式的严重错误
        /// </summary>
        private bool IsProcessCriticalError(Exception ex)
        {
            // RPC 错误
            if (ex.HResult == unchecked((int)0x800706BA) ||
                ex.HResult == unchecked((int)0x800706BE) ||
                ex.Message.Contains("RPC") ||
                ex.Message.Contains("远程过程调用"))
            {
                return true;
            }

            // COM 对象失效
            if (ex is COMException comEx)
            {
                if (comEx.HResult == unchecked((int)0x80010108) ||
                    comEx.HResult == unchecked((int)0x800706BA) ||
                    comEx.HResult == unchecked((int)0x800706BE))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 验证路径长度并在必要时截断文件名
        /// Windows 路径最大长度为 260 字符
        /// </summary>
        private string ValidateAndTruncatePathLength(string filePath)
        {
            const int MAX_PATH_LENGTH = WINDOWS_MAX_PATH_LENGTH;
            const int MIN_FILENAME_LENGTH = ConversionEngine.MIN_FILENAME_LENGTH;

            if (filePath.Length <= MAX_PATH_LENGTH)
            {
                return filePath;
            }

            string directory = Path.GetDirectoryName(filePath);
            string extension = Path.GetExtension(filePath);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);

            // 计算可用于文件名的最大长度
            int availableLength = MAX_PATH_LENGTH - directory.Length - extension.Length - 1; // -1 for directory separator

            if (availableLength < MIN_FILENAME_LENGTH)
            {
                // 如果目录路径太长，无法容纳合理的文件名，抛出异常
                throw new PathTooLongException($"目录路径过长，无法生成有效的文件名: {directory}");
            }

            // 截断文件名并添加哈希值以保证唯一性
            string truncatedName = fileNameWithoutExtension.Substring(0, Math.Min(fileNameWithoutExtension.Length, availableLength - 10));
            string hash = fileNameWithoutExtension.GetHashCode().ToString("X8");
            string newFileName = $"{truncatedName}_{hash}{extension}";

            string newPath = Path.Combine(directory, newFileName);

            Log($"⚠️ 警告: 路径过长已自动截断: {Path.GetFileName(filePath)} → {newFileName}", LogLevel.Warning);

            return newPath;
        }

        /// <summary>
        /// 获取目标文件路径
        /// </summary>
        private string GetToFilePath(string fromRootFolderPath, string toFolderRootPath, string fromFilePath, string toFileName)
        {
            var relativePath = ".";
            if (_viewModel.IsKeepFolderStructure)
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

            // 检查是否是冲突文件
            string targetFileName;
            if (_conflictFiles.ContainsKey(fromFilePath))
            {
                targetFileName = toFileName + ".pdf";
            }
            else
            {
                targetFileName = Path.ChangeExtension(toFileName, ".pdf");
            }

            string fullPath = Path.Combine(toFolderPath, targetFileName);

            // 验证并处理路径长度
            return ValidateAndTruncatePathLength(fullPath);
        }

        /// <summary>
        /// 获取相对路径
        /// </summary>
        private string GetRelativePath(string fromPath, string toPath)
        {
            if (string.IsNullOrEmpty(fromPath)) throw new ArgumentNullException(nameof(fromPath));
            if (string.IsNullOrEmpty(toPath)) throw new ArgumentNullException(nameof(toPath));

            Uri fromUri = new Uri(AppendDirectorySeparatorChar(fromPath));
            Uri toUri = new Uri(AppendDirectorySeparatorChar(toPath));

            if (fromUri.Scheme != toUri.Scheme) { return toPath; }

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
            if (!Path.HasExtension(path) &&
                !path.EndsWith(Path.DirectorySeparatorChar.ToString()))
            {
                return path + Path.DirectorySeparatorChar;
            }
            return path;
        }

        /// <summary>
        /// 处理重复文件
        /// </summary>
        private FileHandleResult HandleDuplicateFile(string originalPath)
        {
            if (string.IsNullOrEmpty(originalPath))
            {
                throw new ArgumentException("文件路径不能为空", nameof(originalPath));
            }

            if (!File.Exists(originalPath))
                return new FileHandleResult
                {
                    FilePath = originalPath,
                    Action = DuplicateFileAction.Rename,
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
                    return new FileHandleResult
                    {
                        FilePath = GetUniqueFilePath(originalPath),
                        Action = DuplicateFileAction.Rename,
                        IsOriginalFile = false
                    };
            }
        }

        /// <summary>
        /// 获取唯一文件路径
        /// </summary>
        private string GetUniqueFilePath(string originalPath)
        {
            if (string.IsNullOrEmpty(originalPath))
            {
                throw new ArgumentException("文件路径不能为空", nameof(originalPath));
            }

            var directory = Path.GetDirectoryName(originalPath);
            var fileNameWithoutExt = Path.GetFileNameWithoutExtension(originalPath);
            var extension = Path.GetExtension(originalPath);

            if (!Directory.Exists(directory))
            {
                throw new DirectoryNotFoundException($"目录不存在: {directory}");
            }

            int counter = 1;
            string newPath;
            const int maxAttempts = MAX_FILENAME_RETRY_ATTEMPTS;

            do
            {
                var newFileName = $"{fileNameWithoutExt} ({counter}){extension}";
                newPath = Path.Combine(directory, newFileName);
                counter++;

                if (counter > maxAttempts)
                {
                    var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    newFileName = $"{fileNameWithoutExt}_{timestamp}{extension}";
                    newPath = Path.Combine(directory, newFileName);
                    break;
                }
            }
            while (File.Exists(newPath));

            return newPath;
        }

        /// <summary>
        /// 统计成功转换的文件数量
        /// </summary>
        private void IncrementSuccessCount(string filePath)
        {
            // 标记文件已成功处理
            _processedSuccessfulFiles.TryAdd(filePath, 0);

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

        /// <summary>
        /// 显示转换结果汇总
        /// </summary>
        public void ShowConversionSummary()
        {
            var failedCount = _failedFiles.Count;
            var skippedCount = _skippedFiles.Count;
            var overwrittenCount = _overwrittenFiles.Count;
            var renamedCount = _renamedFiles.Count;

            var actualSuccessCount = _successfulWordCount + _successfulExcelCount + _successfulPptCount;
            // 修复：始终使用实际成功计数，而不是用总数减去失败数（因为可能存在未统计到的异常）
            var successCount = actualSuccessCount;

            Log($"📊 ============== 转换结果汇总 ==============");

            // 显示总文件数及各类型分布
            var totalDetails = new List<string>();
            if (_totalWordCount > 0) totalDetails.Add($"📄Word {_totalWordCount}");
            if (_totalExcelCount > 0) totalDetails.Add($"📈Excel {_totalExcelCount}");
            if (_totalPptCount > 0) totalDetails.Add($"📽️PPT {_totalPptCount}");

            var totalDetailStr = totalDetails.Count > 0 ? $" | {string.Join(" + ", totalDetails)}" : "";
            Log($"📁 总共文件数：{_totalFilesCount} 个{totalDetailStr}");

            // 显示成功数及各类型分布
            var successDetails = new List<string>();
            if (_successfulWordCount > 0) successDetails.Add($"📄Word {_successfulWordCount}");
            if (_successfulExcelCount > 0) successDetails.Add($"📈Excel {_successfulExcelCount}");
            if (_successfulPptCount > 0) successDetails.Add($"📽️PPT {_successfulPptCount}");

            var successDetailStr = successDetails.Count > 0 ? $" | {string.Join(" + ", successDetails)}" : "";
            Log($"✅ 转换成功：{successCount} 个{successDetailStr}");

            // 显示跳过文件详情
            if (skippedCount > 0)
            {
                var skippedWordCount = CountFilesByExtension(_skippedFiles.Keys, new[] { ".doc", ".docx" });
                var skippedExcelCount = CountFilesByExtension(_skippedFiles.Keys, new[] { ".xls", ".xlsx" });
                var skippedPptCount = CountFilesByExtension(_skippedFiles.Keys, new[] { ".ppt", ".pptx" });

                var skippedDetails = new List<string>();
                if (skippedWordCount > 0) skippedDetails.Add($"📄Word {skippedWordCount}");
                if (skippedExcelCount > 0) skippedDetails.Add($"📈Excel {skippedExcelCount}");
                if (skippedPptCount > 0) skippedDetails.Add($"📽️PPT {skippedPptCount}");

                var skippedDetailStr = skippedDetails.Count > 0 ? $" | {string.Join(" + ", skippedDetails)}" : "";
                Log($"⏭️ 跳过文件：{skippedCount} 个 (目标文件已存在){skippedDetailStr}", LogLevel.Warning);
            }

            // 显示覆盖文件详情
            if (overwrittenCount > 0)
            {
                var overwrittenWordCount = CountFilesByExtension(_overwrittenFiles.Keys, new[] { ".doc", ".docx" });
                var overwrittenExcelCount = CountFilesByExtension(_overwrittenFiles.Keys, new[] { ".xls", ".xlsx" });
                var overwrittenPptCount = CountFilesByExtension(_overwrittenFiles.Keys, new[] { ".ppt", ".pptx" });

                var overwrittenDetails = new List<string>();
                if (overwrittenWordCount > 0) overwrittenDetails.Add($"📄Word {overwrittenWordCount}");
                if (overwrittenExcelCount > 0) overwrittenDetails.Add($"📈Excel {overwrittenExcelCount}");
                if (overwrittenPptCount > 0) overwrittenDetails.Add($"📽️PPT {overwrittenPptCount}");

                var overwrittenDetailStr = overwrittenDetails.Count > 0 ? $" | {string.Join(" + ", overwrittenDetails)}" : "";
                Log($"🔄 覆盖文件：{overwrittenCount} 个 (已覆盖同名目标文件){overwrittenDetailStr}", LogLevel.Warning);
            }

            // 显示重命名文件详情
            if (renamedCount > 0)
            {
                var renamedWordCount = CountFilesByExtension(_renamedFiles.Keys, new[] { ".doc", ".docx" });
                var renamedExcelCount = CountFilesByExtension(_renamedFiles.Keys, new[] { ".xls", ".xlsx" });
                var renamedPptCount = CountFilesByExtension(_renamedFiles.Keys, new[] { ".ppt", ".pptx" });

                var renamedDetails = new List<string>();
                if (renamedWordCount > 0) renamedDetails.Add($"📄Word {renamedWordCount}");
                if (renamedExcelCount > 0) renamedDetails.Add($"📈Excel {renamedExcelCount}");
                if (renamedPptCount > 0) renamedDetails.Add($"📽️PPT {renamedPptCount}");

                var renamedDetailStr = renamedDetails.Count > 0 ? $" | {string.Join(" + ", renamedDetails)}" : "";
                Log($"📝 重命名文件：{renamedCount} 个 (已自动重命名){renamedDetailStr}", LogLevel.Info);
            }

            if (failedCount > 0)
            {
                Log($"❌ 转换失败：{failedCount} 个", LogLevel.Error);
                Log($"💥 失败文件列表：", LogLevel.Error);
                int fileIndex = 1;
                foreach (var failedFile in _failedFiles.Keys)
                {
                    var relativePath = GetRelativePath(_viewModel.FromRootFolderPath, failedFile);
                    Log($"   {fileIndex}. {relativePath}", LogLevel.Error);
                    fileIndex++;
                }
            }

            // 根据转换结果显示相应信息
            if (_wasCancelled)
            {
                Log($"⚠️ 转换被用户取消", LogLevel.Warning);
            }
            else if (failedCount > 0)
            {
                Log($"❌ 部分文件转换失败", LogLevel.Error);
            }
            else if (successCount < _totalFilesCount - skippedCount)
            {
                // 有文件未成功转换（可能因为异常未被捕获）
                var unprocessedCount = _totalFilesCount - skippedCount - successCount;
                Log($"⚠️ 有 {unprocessedCount} 个文件未成功转换（可能因程序异常）", LogLevel.Warning);
            }
            else if (_totalFilesCount > 0)
            {
                Log($"🎉 恭喜！所有文件转换成功！");
            }

            if (_totalFilesCount == 0)
            {
                Log($"⚠ 未找到需要转换的文件", LogLevel.Warning);
            }

            Log($"==========================================");
        }

        private int CountFilesByExtension(IEnumerable<string> files, string[] extensions)
        {
            return files.Count(file =>
            {
                var ext = Path.GetExtension(file).ToLower();
                return extensions.Contains(ext);
            });
        }

        /// <summary>
        /// 设置总文件数（用于统计）
        /// </summary>
        public void SetTotalFilesCount(int wordCount, int excelCount, int pptCount)
        {
            _totalWordCount = wordCount;
            _totalExcelCount = excelCount;
            _totalPptCount = pptCount;
            _totalFilesCount = wordCount + excelCount + pptCount;
        }

        // ==================== 撤回功能相关 ====================

        private void RecordConversionOperation(string targetFilePath, string sourceFilePath, DuplicateFileAction action)
        {
            var operation = new ConversionOperation
            {
                FilePath = targetFilePath,
                SourceFile = sourceFilePath,
                Timestamp = DateTime.Now
            };

            switch (action)
            {
                case DuplicateFileAction.Overwrite:
                    operation.Type = OperationType.OverwriteFile;
                    if (File.Exists(targetFilePath))
                    {
                        operation.BackupPath = CreateBackupFile(targetFilePath);
                        if (string.IsNullOrEmpty(operation.BackupPath))
                        {
                            Log($"警告: 文件 {Path.GetFileName(targetFilePath)} 备份失败，撤回时无法恢复原文件", LogLevel.Warning);
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

        public void RecordDeleteOperation(string deletedFilePath, string backupPath)
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
                var tempDir = Path.Combine(Path.GetTempPath(), "Office2PDF_Backup", _sessionId);

                if (!Directory.Exists(tempDir))
                {
                    Directory.CreateDirectory(tempDir);
                }

                var backupPath = Path.Combine(tempDir, Path.GetFileName(originalFilePath));

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
                Log($"创建备份文件失败: {ex.Message}", LogLevel.Warning);
                return null;
            }
        }

        /// <summary>
        /// 执行撤回操作
        /// </summary>
        public async Task<(int undoneCount, int failedCount)> PerformUndoAsync()
        {
            var undoneCount = 0;
            var failedCount = 0;

            Log("开始撤回操作...", LogLevel.Info);

            var createCount = _conversionHistory.Count(op => op.Type == OperationType.CreateFile);
            var overwriteCount = _conversionHistory.Count(op => op.Type == OperationType.OverwriteFile);
            var deleteCount = _conversionHistory.Count(op => op.Type == OperationType.DeleteFile);
            var dirCount = _conversionHistory.Count(op => op.Type == OperationType.CreateDirectory);

            Log($"将撤回：创建文件 {createCount} 个，覆盖文件 {overwriteCount} 个，删除原文件 {deleteCount} 个，创建目录 {dirCount} 个", LogLevel.Info);

            // 预检查备份文件完整性
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
                Log($"警告: 检测到 {brokenBackups} 个备份文件丢失，对应的覆盖/删除操作无法完全撤回", LogLevel.Warning);
            }

            // 按时间倒序撤回操作 - 将 ConcurrentBag 转换为列表
            var operations = _conversionHistory.ToList();
            for (int i = operations.Count - 1; i >= 0; i--)
            {
                var operation = operations[i];
                try
                {
                    switch (operation.Type)
                    {
                        case OperationType.CreateFile:
                            if (File.Exists(operation.FilePath))
                            {
                                File.Delete(operation.FilePath);
                                Log($"已删除: {Path.GetFileName(operation.FilePath)}");
                                undoneCount++;
                            }
                            break;

                        case OperationType.OverwriteFile:
                            if (File.Exists(operation.FilePath))
                            {
                                File.Delete(operation.FilePath);
                            }
                            if (!string.IsNullOrEmpty(operation.BackupPath) && File.Exists(operation.BackupPath))
                            {
                                File.Copy(operation.BackupPath, operation.FilePath, true);
                                Log($"已恢复: {Path.GetFileName(operation.FilePath)}");
                                undoneCount++;
                            }
                            else
                            {
                                Log($"无法恢复 {Path.GetFileName(operation.FilePath)}: 备份文件不存在", LogLevel.Warning);
                                failedCount++;
                            }
                            break;

                        case OperationType.CreateDirectory:
                            if (Directory.Exists(operation.FilePath))
                            {
                                try
                                {
                                    if (Directory.GetFiles(operation.FilePath).Length == 0 &&
                                        Directory.GetDirectories(operation.FilePath).Length == 0)
                                    {
                                        Directory.Delete(operation.FilePath);
                                        Log($"已删除空目录: {Path.GetFileName(operation.FilePath)}");
                                        undoneCount++;
                                    }
                                    else
                                    {
                                        Log($"目录非空，跳过删除: {Path.GetFileName(operation.FilePath)}", LogLevel.Info);
                                    }
                                }
                                catch (Exception dirEx)
                                {
                                    Log($"删除目录失败: {Path.GetFileName(operation.FilePath)} - {dirEx.Message}", LogLevel.Warning);
                                }
                            }
                            break;

                        case OperationType.DeleteFile:
                            if (!string.IsNullOrEmpty(operation.BackupPath) && File.Exists(operation.BackupPath))
                            {
                                try
                                {
                                    string targetDir = Path.GetDirectoryName(operation.FilePath);
                                    if (!Directory.Exists(targetDir))
                                    {
                                        Directory.CreateDirectory(targetDir);
                                    }

                                    File.Copy(operation.BackupPath, operation.FilePath, true);
                                    Log($"已恢复被删除的文件: {Path.GetFileName(operation.FilePath)}");
                                    undoneCount++;
                                }
                                catch (Exception restoreEx)
                                {
                                    Log($"恢复被删除文件失败: {Path.GetFileName(operation.FilePath)} - {restoreEx.Message}", LogLevel.Error);
                                    failedCount++;
                                }
                            }
                            else
                            {
                                Log($"无法恢复被删除的文件 {Path.GetFileName(operation.FilePath)}: 备份文件不存在", LogLevel.Error);
                                failedCount++;
                            }
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Log($"撤回失败 {Path.GetFileName(operation.FilePath)}: {ex.Message}", LogLevel.Error);
                    failedCount++;
                }

                // 避免UI冻结
                if (i % 10 == 0)
                {
                    await Task.Delay(1);
                }
            }

            // 清理备份文件和历史记录 - 重新创建 ConcurrentBag
            CleanupBackupFiles();
            _conversionHistory = new ConcurrentBag<ConversionOperation>();

            Log($"撤回完成: 成功 {undoneCount} 个，失败 {failedCount} 个",
                failedCount > 0 ? LogLevel.Warning : LogLevel.Info);

            return (undoneCount, failedCount);
        }

        private void CleanupCurrentSessionBackups()
        {
            int deletedFiles = 0;
            int failedFiles = 0;

            try
            {
                foreach (var operation in _conversionHistory)
                {
                    if (!string.IsNullOrEmpty(operation.BackupPath) && File.Exists(operation.BackupPath))
                    {
                        try
                        {
                            File.Delete(operation.BackupPath);
                            deletedFiles++;
                        }
                        catch (Exception ex)
                        {
                            failedFiles++;
                            Log($"清理备份文件失败: {Path.GetFileName(operation.BackupPath)} - {ex.Message}", LogLevel.Warning);
                        }
                    }
                }

                var sessionTempDir = Path.Combine(Path.GetTempPath(), "Office2PDF_Backup", _sessionId);
                if (Directory.Exists(sessionTempDir))
                {
                    try
                    {
                        Directory.Delete(sessionTempDir, true);
                        Log($"✅ 临时目录已清理: {sessionTempDir}");
                    }
                    catch (Exception ex)
                    {
                        Log($"⚠ 清理临时目录失败: {sessionTempDir} - {ex.Message}", LogLevel.Warning);
                    }
                }

                if (deletedFiles > 0)
                {
                    Log($"✅ 已清理 {deletedFiles} 个备份文件{(failedFiles > 0 ? $"，{failedFiles} 个清理失败" : "")}");
                }
            }
            catch (Exception ex)
            {
                Log($"清理备份文件时发生错误: {ex.Message}", LogLevel.Warning);
            }
        }

        private void CleanupBackupFiles()
        {
            try
            {
                CleanupCurrentSessionBackups();
            }
            catch
            {
                // 清理失败不影响主要功能
            }
        }

        // ==================== 删除原文件功能 ====================

        /// <summary>
        /// 删除原文件
        /// </summary>
        public async Task<(int deletedCount, List<string> failedFiles)> DeleteOriginalFilesAsync()
        {
            if (_successfullyConvertedFiles.Count == 0)
                return (0, new List<string>());

            Log($"==============开始删除原文件==============");
            Log($"准备删除文件，释放资源中...");

            await Task.Delay(2000);

            GC.Collect();
            GC.WaitForPendingFinalizers();
            await Task.Delay(500);

            Log($"开始删除文件...");

            var filesToDelete = new List<string>(_successfullyConvertedFiles.Keys);
            _successfullyConvertedFiles.Clear();

            int deletedCount = 0;
            var failedFiles = new List<string>();

            foreach (var filePath in filesToDelete)
            {
                if (!File.Exists(filePath))
                {
                    Log($"✓ 文件已不存在: {Path.GetFileName(filePath)}");
                    continue;
                }

                bool deleted = false;

                for (int attempt = 0; attempt < 5 && !deleted; attempt++)
                {
                    try
                    {
                        if (attempt > 0)
                        {
                            int waitTime = attempt * 1000;
                            Log($"⏳ 第{attempt + 1}次尝试删除: {Path.GetFileName(filePath)} (等待{attempt}秒)");
                            await Task.Delay(waitTime);

                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }

                        File.SetAttributes(filePath, FileAttributes.Normal);

                        string backupPath = null;
                        try
                        {
                            backupPath = CreateBackupFile(filePath);
                        }
                        catch (Exception backupEx)
                        {
                            Log($"⚠ 备份文件失败: {Path.GetFileName(filePath)} - {backupEx.Message}", LogLevel.Warning);
                            continue;
                        }

                        File.Delete(filePath);
                        Log($"✓ 原文件已删除: {Path.GetFileName(filePath)}");

                        RecordDeleteOperation(filePath, backupPath);

                        deletedCount++;
                        deleted = true;
                    }
                    catch (IOException ioEx) when (attempt < 4)
                    {
                        if (attempt == 3)
                        {
                            Log($"⚠ 常规方式删除失败，尝试清理相关进程... (错误: {ioEx.Message})", LogLevel.Warning);
                            await ForceCleanupOfficeProcesses();
                        }
                        else
                        {
                            Log($"⏳ 删除尝试 {attempt + 1} 失败: {Path.GetFileName(filePath)} - {ioEx.Message}", LogLevel.Warning);
                        }
                    }
                    catch (UnauthorizedAccessException uaEx) when (attempt < 4)
                    {
                        Log($"⏳ 删除尝试 {attempt + 1} 失败 (权限不足): {Path.GetFileName(filePath)} - {uaEx.Message}", LogLevel.Warning);
                    }
                    catch (Exception ex)
                    {
                        if (attempt == 4)
                        {
                            Log($"✗ 删除失败 ({ex.GetType().Name}): {Path.GetFileName(filePath)} - {ex.Message}", LogLevel.Warning);
                            failedFiles.Add(filePath);
                        }
                        else
                        {
                            Log($"⏳ 删除尝试 {attempt + 1} 失败 ({ex.GetType().Name}): {Path.GetFileName(filePath)} - {ex.Message}", LogLevel.Warning);
                        }
                    }
                }
            }

            if (failedFiles.Count > 0)
            {
                Log($"⚠ 成功删除 {deletedCount} 个文件，{failedFiles.Count} 个文件删除失败:", LogLevel.Warning);
                foreach (var filePath in failedFiles)
                {
                    Log($"   - {Path.GetFileName(filePath)}", LogLevel.Warning);
                }
                Log($"💡 提示：请手动删除这些文件，或检查文件权限设置。", LogLevel.Info);
            }
            else
            {
                Log($"✅ 成功删除所有 {deletedCount} 个原文件!");
            }

            Log($"==============文件删除完成==============");

            return (deletedCount, failedFiles);
        }

        private async Task ForceCleanupOfficeProcesses()
        {
            try
            {
                Log($"正在清理残留的Office进程...");

                var processNames = new[] {
                    "WINWORD", "EXCEL", "POWERPNT",
                    "wps", "et", "wpp"
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
                                Log($"强制终止残留进程: {processName} (PID: {process.Id})");
                                await Task.Delay(500);
                            }
                        }
                        catch
                        {
                            // 忽略进程终止错误
                        }
                    }
                }

                await Task.Delay(2000);
            }
            catch (Exception ex)
            {
                Log($"清理Office进程时出错: {ex.Message}", LogLevel.Warning);
            }
        }

        // ==================== 辅助方法 ====================

        private void Log(string message, LogLevel level = LogLevel.Trace)
        {
            _logAction?.Invoke(message, level);
        }
    }
}

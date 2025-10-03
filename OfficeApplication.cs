using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

/// <summary>
/// Office引擎检测辅助类
/// </summary>
public static class OfficeEngineHelper
{
    /// <summary>
    /// 检测应用程序路径是否指向WPS Office
    /// </summary>
    public static bool IsWpsEngine(string appPath)
    {
        if (string.IsNullOrEmpty(appPath))
            return false;

        return appPath.IndexOf("king", StringComparison.OrdinalIgnoreCase) >= 0 ||
               appPath.IndexOf("wps", StringComparison.OrdinalIgnoreCase) >= 0;
    }
}

/// <summary>
/// 网络路径处理工具类，专门解决Excel COM组件在UNC网络路径上的兼容性问题
/// </summary>
public static class NetworkPathHelper
{
    /// <summary>
    /// 检查路径是否为UNC网络路径
    /// </summary>
    public static bool IsNetworkPath(string path)
    {
        if (string.IsNullOrEmpty(path))
            return false;

        // 检查UNC路径格式 (\\server\share)
        return path.StartsWith(@"\\") || path.StartsWith("//");
    }

    /// <summary>
    /// 为网络文件创建临时本地副本
    /// </summary>
    public static string CreateLocalTempCopy(string networkFilePath)
    {
        if (!IsNetworkPath(networkFilePath))
            return networkFilePath; // 如果不是网络路径，直接返回原路径

        try
        {
            // 创建基于文件路径哈希的唯一临时文件名
            string fileExtension = Path.GetExtension(networkFilePath);
            string fileName = Path.GetFileNameWithoutExtension(networkFilePath);

            using (var md5 = MD5.Create())
            {
                byte[] hash = md5.ComputeHash(Encoding.UTF8.GetBytes(networkFilePath));
                string hashString = BitConverter.ToString(hash).Replace("-", "").Substring(0, 8);

                string tempDir = Path.Combine(Path.GetTempPath(), "Office2PDF_NetworkTemp");
                if (!Directory.Exists(tempDir))
                    Directory.CreateDirectory(tempDir);

                string tempFilePath = Path.Combine(tempDir, $"{fileName}_{hashString}{fileExtension}");

                // 复制网络文件到本地临时位置
                File.Copy(networkFilePath, tempFilePath, true);
                return tempFilePath;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"无法创建网络文件的本地副本: {networkFilePath}, 错误: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 为网络输出路径创建临时本地路径
    /// </summary>
    public static string CreateLocalTempOutputPath(string networkOutputPath)
    {
        if (!IsNetworkPath(networkOutputPath))
            return networkOutputPath; // 如果不是网络路径，直接返回原路径

        try
        {
            // 创建基于输出路径哈希的唯一临时文件名
            string fileExtension = Path.GetExtension(networkOutputPath);
            string fileName = Path.GetFileNameWithoutExtension(networkOutputPath);

            using (var md5 = MD5.Create())
            {
                byte[] hash = md5.ComputeHash(Encoding.UTF8.GetBytes(networkOutputPath));
                string hashString = BitConverter.ToString(hash).Replace("-", "").Substring(0, 8);

                string tempDir = Path.Combine(Path.GetTempPath(), "Office2PDF_NetworkOutput");
                if (!Directory.Exists(tempDir))
                    Directory.CreateDirectory(tempDir);

                string tempFilePath = Path.Combine(tempDir, $"{fileName}_{hashString}{fileExtension}");
                return tempFilePath;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"无法创建网络输出路径的临时本地路径: {networkOutputPath}, 错误: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 将临时文件复制到网络路径
    /// </summary>
    public static void CopyToNetworkPath(string localTempPath, string networkPath)
    {
        try
        {
            // 确保网络目标目录存在
            string networkDir = Path.GetDirectoryName(networkPath);
            if (!string.IsNullOrEmpty(networkDir) && !Directory.Exists(networkDir))
            {
                Directory.CreateDirectory(networkDir);
            }

            // 复制文件到网络位置
            File.Copy(localTempPath, networkPath, true);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"无法将文件复制到网络路径: {networkPath}, 错误: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 清理临时文件
    /// </summary>
    public static void CleanupTempFile(string tempFilePath)
    {
        try
        {
            if (!string.IsNullOrEmpty(tempFilePath) && File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }
        }
        catch
        {
            // 忽略清理错误
        }
    }

    /// <summary>
    /// 清理所有临时目录（应用程序退出时调用）
    /// </summary>
    public static void CleanupAllTempFiles()
    {
        try
        {
            // 清理输入文件临时目录
            string tempInputDir = Path.Combine(Path.GetTempPath(), "Office2PDF_NetworkTemp");
            if (Directory.Exists(tempInputDir))
            {
                Directory.Delete(tempInputDir, true);
            }

            // 清理输出文件临时目录
            string tempOutputDir = Path.Combine(Path.GetTempPath(), "Office2PDF_NetworkOutput");
            if (Directory.Exists(tempOutputDir))
            {
                Directory.Delete(tempOutputDir, true);
            }
        }
        catch
        {
            // 忽略清理错误
        }
    }
}

public enum OfficeAppType
{
    Word,
    Excel,
    PPT
}

public interface IOfficeApplication : IDisposable
{
    void OpenDocument(string filePath);
    void SaveAsPDF(string toFilePath);
    void CloseDocument();
}

public class MSWordApplication : IOfficeApplication
{
    private Microsoft.Office.Interop.Word.Application _application;
    private Document _document;
    private string _tempFilePath; // 用于跟踪临时文件路径
    private Action<string> _logAction; // 用于输出日志的委托
    private bool _isWpsHijacked = false; // 标记是否被WPS劫持
    public bool IsPrintRevisions { get; set; } = true;

    public MSWordApplication()
    {
        _application = new Microsoft.Office.Interop.Word.Application() { Visible = false, DisplayAlerts = WdAlertLevel.wdAlertsNone };

        // 检测是否被WPS劫持（即使没有日志回调也要检测）
        try
        {
            string appPath = _application.Path;
            _isWpsHijacked = OfficeEngineHelper.IsWpsEngine(appPath);
        }
        catch
        {
            // 检测失败，保持默认值 false
        }
    }

    public MSWordApplication(Action<string> logAction) : this()
    {
        _logAction = logAction;
        try
        {
            // 检测实际连接的Word引擎
            string appPath = _application.Path;
            if (OfficeEngineHelper.IsWpsEngine(appPath))
            {
                _isWpsHijacked = true;
                _logAction?.Invoke($"💡 提示: 后台连接到 WPS 文字，路径为 {appPath}");
            }
            else
            {
                _logAction?.Invoke($"💡 提示: 后台连接到 Microsoft Word，路径为 {appPath}");
            }
        }
        catch (Exception ex)
        {
            _logAction?.Invoke($"检测Microsoft Word状态失败: {ex.Message}");
        }
    }

    public void OpenDocument(string filePath)
    {
        // 如果是网络路径，创建本地临时副本
        if (NetworkPathHelper.IsNetworkPath(filePath))
        {
            _tempFilePath = NetworkPathHelper.CreateLocalTempCopy(filePath);
            _document = _application.Documents.Open(_tempFilePath, ReadOnly: true);
        }
        else
        {
            _document = _application.Documents.Open(filePath, ReadOnly: true);
        }
    }

    public void SaveAsPDF(string toFilePath)
    {
        if (_document != null)
        {
            try
            {
                var directory = Path.GetDirectoryName(toFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // 只有在被WPS劫持的情况下才使用预处理方法
                if (_isWpsHijacked)
                {
                    if (IsPrintRevisions)
                    {
                        // 如果勾选了“打印批注”，则直接导出，保留所有标记
                        _document.SaveAs2(toFilePath, WdSaveFormat.wdFormatPDF);
                    }
                    else
                    {
                        // 如果未勾选，才调用预处理方法来生成干净的PDF
                        SaveAsPDFWithPreprocessing(toFilePath);
                    }
                }
                else
                {
                    if (IsPrintRevisions)
                    {
                        // 直接导出，显示所有批注和修订
                        var originalShowRevisions = _document.ShowRevisions;
                        _document.ShowRevisions = true;
                        _document.SaveAs2(toFilePath, WdSaveFormat.wdFormatPDF);
                        _document.ShowRevisions = originalShowRevisions;
                    }
                    else
                    {
                        // 隐藏批注和修订后导出
                        var originalShowRevisions = _document.ShowRevisions;
                        _document.ShowRevisions = false;
                        _document.SaveAs2(toFilePath, WdSaveFormat.wdFormatPDF);
                        _document.ShowRevisions = originalShowRevisions;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"无法保存PDF: {toFilePath}, 错误: {ex.Message}", ex);
            }
        }
    }

    /// <summary>
    /// 使用文档预处理的方式保存PDF：创建临时副本，物理删除批注和接受修订
    /// </summary>
    private void SaveAsPDFWithPreprocessing(string toFilePath)
    {
        string tempDocPath = null;
        Document tempDoc = null;

        try
        {
            // 1. 创建临时文件路径
            tempDocPath = Path.Combine(Path.GetTempPath(), $"Office2PDF_Temp_{Guid.NewGuid():N}.docx");

            // 2. 保存当前文档为临时副本
            _document.SaveAs2(tempDocPath);

            // 3. 打开临时副本进行预处理
            tempDoc = _application.Documents.Open(tempDocPath, ReadOnly: false);

            // 4. 删除所有批注
            try
            {
                var comments = tempDoc.Comments;
                while (comments.Count > 0)
                {
                    comments[1].Delete();
                }
            }
            catch (Exception ex)
            {
                // 批注删除失败时记录但不中断处理
                _logAction?.Invoke($"⚠️ MS Word: 批注删除失败，PDF可能保留批注 - {ex.Message}");
            }

            // 5. 接受所有修订标记
            try
            {
                var revisions = tempDoc.Revisions;
                if (revisions.Count > 0)
                {
                    revisions.AcceptAll();
                }
            }
            catch (Exception ex)
            {
                // 修订处理失败时记录但不中断处理
                _logAction?.Invoke($"⚠️ MS Word: 修订接受失败，PDF可能保留修订标记 - {ex.Message}");
            }

            // 6. 保存临时文档的更改
            tempDoc.Save();

            // 7. 导出预处理后的文档为PDF
            tempDoc.SaveAs2(toFilePath, WdSaveFormat.wdFormatPDF);
        }
        finally
        {
            // 8. 清理临时资源
            if (tempDoc != null)
            {
                try
                {
                    tempDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
                    Marshal.ReleaseComObject(tempDoc);
                }
                catch { /* ignore cleanup errors */ }
            }

            // 9. 删除临时文件
            if (!string.IsNullOrEmpty(tempDocPath) && File.Exists(tempDocPath))
            {
                try
                {
                    File.Delete(tempDocPath);
                }
                catch { /* ignore temp file deletion errors */ }
            }
        }
    }

    public void CloseDocument()
    {
        if (_document != null)
        {
            _document.Close(WdSaveOptions.wdDoNotSaveChanges);
            try { Marshal.ReleaseComObject(_document); } catch { /* COM清理失败不影响程序继续 */ }
            _document = null;
        }
    }

    public void Dispose()
    {
        _application.Quit();
        try { Marshal.ReleaseComObject(_application); } catch { /* COM清理失败不影响程序继续 */ }

        // 清理临时文件
        NetworkPathHelper.CleanupTempFile(_tempFilePath);
    }
}

public class MSExcelApplication : IOfficeApplication
{
    private Microsoft.Office.Interop.Excel.Application _application;
    private Workbook _workbook;
    private string _tempFilePath; // 用于跟踪临时文件路径
    private Action<string> _logAction; // 用于输出日志的委托
    public bool IsConvertOneSheetOnePDF { get; set; } = true;

    public MSExcelApplication()
    {
        _application = new Microsoft.Office.Interop.Excel.Application() { Visible = false, DisplayAlerts = false };
    }

    public MSExcelApplication(Action<string> logAction) : this()
    {
        _logAction = logAction;
        try
        {
            // 检测实际连接的Excel引擎
            string appPath = _application.Path;
            if (OfficeEngineHelper.IsWpsEngine(appPath))
            {
                _logAction?.Invoke($"💡 提示: 后台连接到 WPS 表格，路径为 {appPath}");
            }
            else
            {
                _logAction?.Invoke($"💡 提示: 后台连接到 Microsoft Excel，路径为 {appPath}");
            }
        }
        catch (Exception ex)
        {
            _logAction?.Invoke($"检测Microsoft Excel状态失败: {ex.Message}");
        }
    }

    public void OpenDocument(string filePath)
    {
        // 如果是网络路径，创建本地临时副本
        if (NetworkPathHelper.IsNetworkPath(filePath))
        {
            _tempFilePath = NetworkPathHelper.CreateLocalTempCopy(filePath);
            _workbook = _application.Workbooks.Open(_tempFilePath, ReadOnly: true);
        }
        else
        {
            _workbook = _application.Workbooks.Open(filePath, ReadOnly: true);
        }
    }

    public void SaveAsPDF(string toFilePath)
    {
        if (_workbook != null)
        {
            // 检查目标路径是否为网络路径
            bool isNetworkOutput = NetworkPathHelper.IsNetworkPath(toFilePath);
            string actualOutputPath = isNetworkOutput ? NetworkPathHelper.CreateLocalTempOutputPath(toFilePath) : toFilePath;

            try
            {
                // 确保输出目录存在
                var directory = Path.GetDirectoryName(actualOutputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                if (IsConvertOneSheetOnePDF)
                {
                    // 每个sheet保存为单独的PDF（跳过隐藏的Sheet）
                    if (_workbook.Sheets.Count == 1)
                    {
                        var worksheet = _workbook.Sheets[1];

                        // 检查 Sheet 是否可见（使用 int 转换以兼容 MS Office 和 WPS）
                        // xlSheetVisible = -1 (可见)
                        int visibilityStatus = (int)worksheet.Visible;
                        if (visibilityStatus == -1)
                        {
                            worksheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, actualOutputPath);

                            // 如果是网络输出，复制到最终位置
                            if (isNetworkOutput)
                            {
                                NetworkPathHelper.CopyToNetworkPath(actualOutputPath, toFilePath);
                                NetworkPathHelper.CleanupTempFile(actualOutputPath);
                            }
                        }

                        try { Marshal.ReleaseComObject(worksheet); } catch { /* COM清理失败不影响程序继续 */ }
                    }
                    else
                    {
                        for (int i = 1; i <= _workbook.Sheets.Count; i++)
                        {
                            var worksheet = _workbook.Sheets[i];
                            if (directory == null) throw new ArgumentException("文件没有目录", nameof(toFilePath));

                            // 检查 Sheet 是否可见，跳过隐藏的 Sheet
                            // xlSheetVisible = -1 (可见), xlSheetHidden = 0 (隐藏), xlSheetVeryHidden = 2 (非常隐藏)
                            int visibilityStatus = (int)worksheet.Visible;
                            if (visibilityStatus != -1)
                            {
                                try { Marshal.ReleaseComObject(worksheet); } catch { /* COM清理失败不影响程序继续 */ }
                                continue;
                            }

                            // 获取 Sheet 的实际名称并处理特殊字符
                            string sheetName = worksheet.Name;
                            string safeSheetName = string.Join("_", sheetName.Split(Path.GetInvalidFileNameChars()));

                            // 在本地（或临时）目录生成sheet PDF，使用实际的 Sheet 名称
                            string sheetOutputPath = Path.Combine(directory, $"{Path.GetFileNameWithoutExtension(actualOutputPath)}_{safeSheetName}.pdf");

                            worksheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, sheetOutputPath);
                            try { Marshal.ReleaseComObject(worksheet); } catch { /* COM清理失败不影响程序继续 */ }

                            // 如果是网络输出，复制到最终位置
                            if (isNetworkOutput)
                            {
                                string finalDirectory = Path.GetDirectoryName(toFilePath);
                                string finalSheetPath = Path.Combine(finalDirectory, $"{Path.GetFileNameWithoutExtension(toFilePath)}_{safeSheetName}.pdf");
                                NetworkPathHelper.CopyToNetworkPath(sheetOutputPath, finalSheetPath);
                                NetworkPathHelper.CleanupTempFile(sheetOutputPath);
                            }
                        }
                    }
                }
                else
                {
                    // 所有sheet保存为一个PDF
                    _workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, actualOutputPath);

                    // 如果是网络输出，复制到最终位置
                    if (isNetworkOutput)
                    {
                        NetworkPathHelper.CopyToNetworkPath(actualOutputPath, toFilePath);
                        NetworkPathHelper.CleanupTempFile(actualOutputPath);
                    }
                }
                // 触发垃圾回收，强制释放所有sheet的COM对象，防止进程残留或处理下一个文件时卡顿
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception)
            {
                // 清理临时文件
                if (isNetworkOutput && !string.IsNullOrEmpty(actualOutputPath) && File.Exists(actualOutputPath))
                {
                    NetworkPathHelper.CleanupTempFile(actualOutputPath);
                }
                throw;
            }
        }
    }

    public void CloseDocument()
    {
        if (_workbook != null)
        {
            _workbook.Close(false);
            try { Marshal.ReleaseComObject(_workbook); } catch { /* COM清理失败不影响程序继续 */ }
            _workbook = null;
        }

        // 清理临时文件
        if (!string.IsNullOrEmpty(_tempFilePath))
        {
            NetworkPathHelper.CleanupTempFile(_tempFilePath);
            _tempFilePath = null;
        }
    }

    public void Dispose()
    {
        CloseDocument(); // 确保临时文件被清理
        _application.Quit();
        try { Marshal.ReleaseComObject(_application); } catch { /* COM清理失败不影响程序继续 */ }
    }
}

public class MSPowerPointApplication : IOfficeApplication
{
    private Microsoft.Office.Interop.PowerPoint.Application _application;
    private Presentation _presentation;
    private string _tempFilePath; // 用于跟踪临时文件路径
    private Action<string> _logAction; // 用于输出日志的委托

    public MSPowerPointApplication()
    {
        _application = new Microsoft.Office.Interop.PowerPoint.Application();
    }

    public MSPowerPointApplication(Action<string> logAction) : this()
    {
        _logAction = logAction;
        try
        {
            // 检测实际连接的PowerPoint引擎
            string appPath = _application.Path;
            if (OfficeEngineHelper.IsWpsEngine(appPath))
            {
                _logAction?.Invoke($"💡 提示: 后台连接到 WPS 演示，路径为 {appPath}");
            }
            else
            {
                _logAction?.Invoke($"💡 提示: 后台连接到 Microsoft PowerPoint，路径为 {appPath}");
            }
        }
        catch (Exception ex)
        {
            _logAction?.Invoke($"检测Microsoft PowerPoint状态失败: {ex.Message}");
        }
    }

    public void OpenDocument(string filePath)
    {
        // 如果是网络路径，创建本地临时副本
        if (NetworkPathHelper.IsNetworkPath(filePath))
        {
            _tempFilePath = NetworkPathHelper.CreateLocalTempCopy(filePath);
            _presentation = _application.Presentations.Open(_tempFilePath, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);
        }
        else
        {
            _presentation = _application.Presentations.Open(filePath, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);
        }
    }

    public void SaveAsPDF(string toFilePath)
    {
        if (_presentation != null)
        {
            // 检查目标路径是否为网络路径
            bool isNetworkOutput = NetworkPathHelper.IsNetworkPath(toFilePath);
            string actualOutputPath = isNetworkOutput ? NetworkPathHelper.CreateLocalTempOutputPath(toFilePath) : toFilePath;

            try
            {
                // 确保输出目录存在
                var directory = Path.GetDirectoryName(actualOutputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                _presentation.ExportAsFixedFormat(actualOutputPath, PpFixedFormatType.ppFixedFormatTypePDF);

                // 如果是网络输出，复制到最终位置
                if (isNetworkOutput)
                {
                    NetworkPathHelper.CopyToNetworkPath(actualOutputPath, toFilePath);
                    NetworkPathHelper.CleanupTempFile(actualOutputPath);
                }
            }
            catch (Exception)
            {
                // 清理临时文件
                if (isNetworkOutput && !string.IsNullOrEmpty(actualOutputPath) && File.Exists(actualOutputPath))
                {
                    NetworkPathHelper.CleanupTempFile(actualOutputPath);
                }
                throw;
            }
        }
    }

    public void CloseDocument()
    {
        if (_presentation != null)
        {
            _presentation.Close();
            try { Marshal.ReleaseComObject(_presentation); } catch { /* COM清理失败不影响程序继续 */ }
            _presentation = null;
        }
    }

    public void Dispose()
    {
        _application.Quit();
        try { Marshal.ReleaseComObject(_application); } catch { /* COM清理失败不影响程序继续 */ }

        // 清理临时文件
        NetworkPathHelper.CleanupTempFile(_tempFilePath);
    }
}

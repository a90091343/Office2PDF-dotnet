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
        _document = _application.Documents.Open(filePath, ReadOnly: true);
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
    }
}

public class MSExcelApplication : IOfficeApplication
{
    private Microsoft.Office.Interop.Excel.Application _application;
    private Workbook _workbook;
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
        _workbook = _application.Workbooks.Open(filePath, ReadOnly: true);
    }

    public void SaveAsPDF(string toFilePath)
    {
        if (_workbook != null)
        {
            try
            {
                // 确保输出目录存在
                var directory = Path.GetDirectoryName(toFilePath);
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
                            worksheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, toFilePath);
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

                            // 生成每个 sheet 的 PDF 文件，使用实际的 Sheet 名称
                            string sheetOutputPath = Path.Combine(directory, $"{Path.GetFileNameWithoutExtension(toFilePath)}_{safeSheetName}.pdf");

                            worksheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, sheetOutputPath);
                            try { Marshal.ReleaseComObject(worksheet); } catch { /* COM清理失败不影响程序继续 */ }
                        }
                    }
                }
                else
                {
                    // 所有sheet保存为一个PDF
                    _workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, toFilePath);
                }

                // 触发垃圾回收，强制释放所有sheet的COM对象，防止进程残留或处理下一个文件时卡顿
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"无法保存PDF: {toFilePath}, 错误: {ex.Message}", ex);
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
    }

    public void Dispose()
    {
        CloseDocument();
        _application.Quit();
        try { Marshal.ReleaseComObject(_application); } catch { /* COM清理失败不影响程序继续 */ }
    }
}

public class MSPowerPointApplication : IOfficeApplication
{
    private Microsoft.Office.Interop.PowerPoint.Application _application;
    private Presentation _presentation;
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
        _presentation = _application.Presentations.Open(filePath, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);
    }

    public void SaveAsPDF(string toFilePath)
    {
        if (_presentation != null)
        {
            try
            {
                // 确保输出目录存在
                var directory = Path.GetDirectoryName(toFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                _presentation.ExportAsFixedFormat(toFilePath, PpFixedFormatType.ppFixedFormatTypePDF);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"无法保存PDF: {toFilePath}, 错误: {ex.Message}", ex);
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
    }
}

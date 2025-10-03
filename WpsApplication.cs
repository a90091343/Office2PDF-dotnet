using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

// WPS Office COM Interop类
// WPS Writer (Word)
public class WpsWriterApplication : IOfficeApplication
{
    private dynamic _application;
    private dynamic _document;
    private Action<string> _logAction; // 用于输出日志的委托
    public bool IsPrintRevisions { get; set; } = true;

    public WpsWriterApplication()
    {
        try
        {
            // Use the correct ProgID with uppercase K as shown in standard implementation
            _application = Activator.CreateInstance(Type.GetTypeFromProgID("KWps.Application"));
            _application.Visible = false;
            // Ignore warning prompts - very important!
            _application.DisplayAlerts = false;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("无法启动WPS Writer，请确保已安装WPS Office并使用 'KWps.Application' ProgID", ex);
        }
    }

    public WpsWriterApplication(Action<string> logAction) : this()
    {
        _logAction = logAction;
    }

    public void OpenDocument(string filePath)
    {
        try
        {
            // Use the same pattern as standard implementation with Visible parameter
            _document = _application.Documents.Open(filePath, Visible: false);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"无法打开文档: {filePath}", ex);
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

                if (IsPrintRevisions)
                {
                    // 直接导出，显示所有批注和修订
                    _document.ExportAsFixedFormat(toFilePath, 17); // 17 = wdExportFormatPDF
                }
                else
                {
                    // 使用预处理方法：创建临时副本，删除批注和修订
                    SaveAsPDFWithPreprocessing(toFilePath);
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
        dynamic tempDoc = null;

        try
        {
            // 1. 创建临时文件路径
            tempDocPath = Path.Combine(Path.GetTempPath(), $"Office2PDF_Temp_{Guid.NewGuid():N}.docx");

            // 2. 保存当前文档为临时副本
            _document.SaveAs2(tempDocPath);

            // 3. 打开临时副本进行预处理
            tempDoc = _application.Documents.Open(tempDocPath, Visible: false);

            // 4. 删除所有批注
            try
            {
                var comments = tempDoc.Comments;
                while (comments.Count > 0)
                {
                    comments.Item(1).Delete();
                }
            }
            catch (Exception ex)
            {
                // 批注删除失败时记录但不中断处理
                _logAction?.Invoke($"⚠️ WPS Writer: 批注删除失败，PDF可能保留批注 - {ex.Message}");
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
                _logAction?.Invoke($"⚠️ WPS Writer: 修订接受失败，PDF可能保留修订标记 - {ex.Message}");
            }

            // 6. 保存临时文档的更改
            tempDoc.Save();

            // 7. 导出预处理后的文档为PDF
            tempDoc.ExportAsFixedFormat(toFilePath, 17); // 17 = wdExportFormatPDF
        }
        finally
        {
            // 8. 清理临时资源
            if (tempDoc != null)
            {
                try
                {
                    tempDoc.Close(false);
                    if (Marshal.IsComObject(tempDoc))
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
            try
            {
                _document.Close(false);
                if (Marshal.IsComObject(_document))
                    Marshal.ReleaseComObject(_document);
                _document = null;
            }
            catch { /* COM清理失败不影响程序继续 */ }
        }
    }

    public void Dispose()
    {
        try
        {
            CloseDocument();
            if (_application != null)
            {
                _application.Quit();
                if (Marshal.IsComObject(_application))
                    Marshal.ReleaseComObject(_application);
                _application = null;
            }
        }
        catch { /* COM清理失败不影响程序继续 */ }
    }
}

// WPS Spreadsheets (Excel)
public class WpsSpreadsheetApplication : IOfficeApplication
{
    private dynamic _application;
    private dynamic _workbook;
    private string _tempFilePath; // 用于跟踪临时文件路径
    public bool IsConvertOneSheetOnePDF { get; set; } = true;

    public WpsSpreadsheetApplication()
    {
        try
        {
            // Use the correct ProgID with all uppercase KET as shown in standard implementation
            _application = Activator.CreateInstance(Type.GetTypeFromProgID("KET.Application"));
            _application.Visible = false;
            _application.DisplayAlerts = false;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("无法启动WPS Spreadsheets，请确保已安装WPS Office并使用 'KET.Application' ProgID", ex);
        }
    }

    public void OpenDocument(string filePath)
    {
        try
        {
            // 如果是网络路径，创建本地临时副本
            if (NetworkPathHelper.IsNetworkPath(filePath))
            {
                _tempFilePath = NetworkPathHelper.CreateLocalTempCopy(filePath);
                // Use the same pattern as standard implementation with missing parameters
                object missing = Type.Missing;
                _workbook = _application.Workbooks.Open(_tempFilePath, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            }
            else
            {
                // Use the same pattern as standard implementation with missing parameters
                object missing = Type.Missing;
                _workbook = _application.Workbooks.Open(filePath, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"无法打开工作簿: {filePath}", ex);
        }
    }

    public void SaveAsPDF(string toFilePath)
    {
        if (_workbook != null)
        {
            try
            {
                // 检查目标路径是否为网络路径
                bool isNetworkOutput = NetworkPathHelper.IsNetworkPath(toFilePath);
                string actualOutputPath = isNetworkOutput ? NetworkPathHelper.CreateLocalTempOutputPath(toFilePath) : toFilePath;

                var directory = Path.GetDirectoryName(actualOutputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                if (IsConvertOneSheetOnePDF)
                {
                    string fileExt = Path.GetExtension(actualOutputPath);
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(actualOutputPath);

                    foreach (var sheetObj in _workbook.Worksheets)
                    {
                        dynamic sheet = null;
                        string sheetTempPath = null;
                        string finalSheetPath = null;
                        try
                        {
                            sheet = sheetObj; // dynamic cast

                            // 检查 Sheet 是否可见（跳过隐藏的 Sheet）
                            // WPS: Visible 属性值：-1 = 可见，0 = 隐藏，2 = 非常隐藏
                            int visibilityStatus = sheet.Visible;
                            if (visibilityStatus != -1)
                            {
                                continue; // 跳过隐藏的 Sheet
                            }

                            string sheetName = sheet.Name;
                            string safeSheetName = string.Join("_", sheetName.Split(Path.GetInvalidFileNameChars()));
                            sheetTempPath = Path.Combine(directory, $"{fileNameWithoutExt}_{safeSheetName}{fileExt}");
                            object missing = Type.Missing;
                            sheet.ExportAsFixedFormat(0, sheetTempPath, 0, true, false, missing, missing, missing, missing);

                            // 如果是网络输出，复制到最终位置
                            if (isNetworkOutput)
                            {
                                string finalDirectory = Path.GetDirectoryName(toFilePath);
                                finalSheetPath = Path.Combine(finalDirectory, $"{Path.GetFileNameWithoutExtension(toFilePath)}_{safeSheetName}{Path.GetExtension(toFilePath)}");
                                NetworkPathHelper.CopyToNetworkPath(sheetTempPath, finalSheetPath);
                                NetworkPathHelper.CleanupTempFile(sheetTempPath);
                            }
                        }
                        catch (Exception)
                        {
                            // 清理临时文件
                            if (isNetworkOutput && !string.IsNullOrEmpty(sheetTempPath))
                            {
                                NetworkPathHelper.CleanupTempFile(sheetTempPath);
                            }
                            // 记录导出sheet失败的警告
#if DEBUG
                            // System.Diagnostics.Debug.WriteLine($"[WPS] Sheet导出失败: {sheet?.Name}, 错误: {ex.Message}");
#endif
                        }
                        finally
                        {
                            if (sheet != null && Marshal.IsComObject(sheet))
                            {
                                try { Marshal.ReleaseComObject(sheet); } catch { /* COM清理失败不影响程序继续 */ }
                            }
                            sheet = null;
                        }
                    }
                    // 触发GC以加速释放（大量Sheet时尤为重要）
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                else
                {
                    // Use the same pattern as standard implementation for full workbook export
                    object missing = Type.Missing;
                    _workbook.ExportAsFixedFormat(0, actualOutputPath, 0, true, false, missing, missing, missing, missing);

                    // 如果是网络输出，复制到最终位置
                    if (isNetworkOutput)
                    {
                        NetworkPathHelper.CopyToNetworkPath(actualOutputPath, toFilePath);
                        NetworkPathHelper.CleanupTempFile(actualOutputPath);
                    }
                }
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
            try
            {
                _workbook.Close(false);
                if (Marshal.IsComObject(_workbook))
                    Marshal.ReleaseComObject(_workbook);
                _workbook = null;
            }
            catch { /* COM清理失败不影响程序继续 */ }
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
        try
        {
            CloseDocument();
            if (_application != null)
            {
                _application.Quit();
                if (Marshal.IsComObject(_application))
                    Marshal.ReleaseComObject(_application);
                _application = null;
            }
        }
        catch { /* COM清理失败不影响程序继续 */ }
    }
}

// WPS Presentation (PowerPoint)
public class WpsPresentationApplication : IOfficeApplication
{
    private dynamic _application;
    private dynamic _presentation;
    private string _tempFilePath; // 用于跟踪临时文件路径

    public WpsPresentationApplication()
    {
        try
        {
            // Use the correct WPS Presentation ProgID
            Type type = Type.GetTypeFromProgID("KWPP.Application");
            _application = Activator.CreateInstance(type);

            // Ignore warning prompts - very important!
            _application.DisplayAlerts = false;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("无法启动WPS Presentation，请确保已安装WPS Office并使用 'KWPP.Application' ProgID", ex);
        }
    }

    public void OpenDocument(string filePath)
    {
        try
        {
            // 如果是网络路径，创建本地临时副本
            if (NetworkPathHelper.IsNetworkPath(filePath))
            {
                _tempFilePath = NetworkPathHelper.CreateLocalTempCopy(filePath);
                // Use the exact same pattern as standard implementation
                // MsoTriState.msoCTrue = -1 (equivalent to the reference implementation)
                _presentation = _application.Presentations.Open(_tempFilePath, -1, -1, -1);
            }
            else
            {
                // Use the exact same pattern as standard implementation
                // MsoTriState.msoCTrue = -1 (equivalent to the reference implementation)
                _presentation = _application.Presentations.Open(filePath, -1, -1, -1);
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"无法打开演示文稿: {filePath}", ex);
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
                var directory = Path.GetDirectoryName(actualOutputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Use SaveAs method exactly as shown in the standard reference code
                // PpSaveAsFileType.ppSaveAsPDF = 32, MsoTriState.msoTrue = -1
                _presentation.SaveAs(actualOutputPath, 32, -1);

                // 如果是网络输出，复制到最终位置
                if (isNetworkOutput)
                {
                    NetworkPathHelper.CopyToNetworkPath(actualOutputPath, toFilePath);
                    NetworkPathHelper.CleanupTempFile(actualOutputPath);
                }
            }
            catch (Exception ex)
            {
                // 清理临时文件
                if (isNetworkOutput && !string.IsNullOrEmpty(actualOutputPath) && File.Exists(actualOutputPath))
                {
                    NetworkPathHelper.CleanupTempFile(actualOutputPath);
                }
                throw new InvalidOperationException($"无法保存PDF: {toFilePath}, 错误: {ex.Message}", ex);
            }
        }
    }

    public void CloseDocument()
    {
        if (_presentation != null)
        {
            try
            {
                // Use the presentation close method
                _presentation.Close();
                if (Marshal.IsComObject(_presentation))
                    Marshal.ReleaseComObject(_presentation);
                _presentation = null;
            }
            catch { /* COM清理失败不影响程序继续 */ }
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
        try
        {
            CloseDocument();
            if (_application != null)
            {
                _application.Quit();
                if (Marshal.IsComObject(_application))
                    Marshal.ReleaseComObject(_application);
                _application = null;
            }
        }
        catch { /* COM清理失败不影响程序继续 */ }
    }
}

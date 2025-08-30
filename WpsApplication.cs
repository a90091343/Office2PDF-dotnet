using System;
using System.IO;
using System.Runtime.InteropServices;

// WPS Office COM Interop类
// WPS Writer (Word)
public class WpsWriterApplication : IOfficeApplication
{
    private dynamic _application;
    private dynamic _document;
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

                if (!IsPrintRevisions)
                {
                    try { _document.ShowRevisions = false; } catch { /* ignore */ }
                }

                _document.ExportAsFixedFormat(toFilePath, 17); // 17 = wdExportFormatPDF
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"无法保存PDF: {toFilePath}, 错误: {ex.Message}", ex);
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
            catch { }
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
        catch { }
    }
}

// WPS Spreadsheets (Excel)
public class WpsSpreadsheetApplication : IOfficeApplication
{
    private dynamic _application;
    private dynamic _workbook;
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
            // Use the same pattern as standard implementation with missing parameters
            object missing = Type.Missing;
            _workbook = _application.Workbooks.Open(filePath, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
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
                var directory = Path.GetDirectoryName(toFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                if (IsConvertOneSheetOnePDF)
                {
                    string fileExt = Path.GetExtension(toFilePath);
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(toFilePath);

                    foreach (var sheetObj in _workbook.Worksheets)
                    {
                        dynamic sheet = null;
                        try
                        {
                            sheet = sheetObj; // dynamic cast
                            string sheetName = sheet.Name;
                            string safeSheetName = string.Join("_", sheetName.Split(Path.GetInvalidFileNameChars()));
                            string singleSheetPdfPath = Path.Combine(directory, $"{fileNameWithoutExt}_{safeSheetName}{fileExt}");
                            object missing = Type.Missing;
                            sheet.ExportAsFixedFormat(0, singleSheetPdfPath, 0, true, false, missing, missing, missing, missing);
                        }
                        catch (Exception)
                        {
                            try
                            {
                                // 静默处理工作表导出错误
                            }
                            catch { }
                        }
                        finally
                        {
                            if (sheet != null && Marshal.IsComObject(sheet))
                            {
                                try { Marshal.ReleaseComObject(sheet); } catch { }
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
                    _workbook.ExportAsFixedFormat(0, toFilePath, 0, true, false, missing, missing, missing, missing);
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
            catch { }
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
        catch { }
    }
}

// WPS Presentation (PowerPoint)
public class WpsPresentationApplication : IOfficeApplication
{
    private dynamic _application;
    private dynamic _presentation;

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
            // Use the exact same pattern as standard implementation
            // MsoTriState.msoCTrue = -1 (equivalent to the reference implementation)
            _presentation = _application.Presentations.Open(filePath, -1, -1, -1);
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
            try
            {
                var directory = Path.GetDirectoryName(toFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Use SaveAs method exactly as shown in the standard reference code
                // PpSaveAsFileType.ppSaveAsPDF = 32, MsoTriState.msoTrue = -1
                _presentation.SaveAs(toFilePath, 32, -1);
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
            try
            {
                // Use the presentation close method
                _presentation.Close();
                if (Marshal.IsComObject(_presentation))
                    Marshal.ReleaseComObject(_presentation);
                _presentation = null;
            }
            catch { }
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
        catch { }
    }
}

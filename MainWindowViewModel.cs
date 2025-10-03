using System;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Data;

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

    // 日志级别
    public enum LogLevel
    {
        Trace,
        Info,
        Warning,
        Error
    }

    public class MainWindowViewModel : INotifyPropertyChanged
    {
        public MainWindowViewModel()
        {
            // 初始化时根据各个转换类型的状态更新"全选"状态
            UpdateIsConvertAll();
        }

        // 版本号属性，从Assembly中提取
        public string VersionNumber
        {
            get
            {
                var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                return $"{version.Major}.{version.Minor}.{version.Build}";
            }
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

        private bool _isConvertExcel = true;  // 默认勾选Excel
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

        private bool _useWpsOffice = false;  // 默认选择自动（推荐）
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

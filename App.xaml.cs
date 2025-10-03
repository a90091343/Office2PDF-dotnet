using System.Windows;

namespace Office2PDF
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public static string CommandLineFolder { get; private set; }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // 检查是否有命令行参数（拖拽文件夹到 exe）
            if (e.Args.Length > 0)
            {
                string path = e.Args[0];

                // 检查路径是否存在且是文件夹
                if (System.IO.Directory.Exists(path))
                {
                    // 保存到静态属性，MainWindow 启动后会读取
                    CommandLineFolder = path;
                }
            }
        }
    }
}

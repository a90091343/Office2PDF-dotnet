using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

/// <summary>
/// 网络路径处理装饰器：透明地处理网络路径，让所有 Office 应用类无需关心网络路径的特殊处理
/// 使用目录级别的映射策略，支持单文件和多文件输出场景（如 Excel 多 sheet）
/// </summary>
public class NetworkPathHandlingDecorator : IOfficeApplication
{
    private readonly IOfficeApplication _innerApplication;
    private string _tempInputFilePath;      // 输入文件的临时本地副本路径
    private string _originalInputPath;       // 原始输入路径（保存用于日志）

    private string _networkOutputDirectory;  // 原始网络输出目录
    private string _localTempOutputDir;      // 本地临时输出目录（目录级映射）
    private bool _isOutputToNetwork;         // 标记输出是否到网络路径

    /// <summary>
    /// 获取被包装的内部应用实例（用于访问具体应用类的特定属性）
    /// </summary>
    public IOfficeApplication InnerApplication => _innerApplication;

    public NetworkPathHandlingDecorator(IOfficeApplication innerApplication)
    {
        _innerApplication = innerApplication ?? throw new ArgumentNullException(nameof(innerApplication));
    }

    public void OpenDocument(string filePath)
    {
        _originalInputPath = filePath;

        // 处理输入网络路径：创建本地临时副本
        if (NetworkPathHelper.IsNetworkPath(filePath))
        {
            _tempInputFilePath = NetworkPathHelper.CreateLocalTempCopy(filePath);
            _innerApplication.OpenDocument(_tempInputFilePath);
        }
        else
        {
            _innerApplication.OpenDocument(filePath);
        }
    }

    public void SaveAsPDF(string toFilePath)
    {
        _isOutputToNetwork = NetworkPathHelper.IsNetworkPath(toFilePath);

        if (_isOutputToNetwork)
        {
            // 网络路径输出：使用目录级别映射
            _networkOutputDirectory = Path.GetDirectoryName(toFilePath);
            string fileName = Path.GetFileName(toFilePath);

            // 创建本地临时输出目录
            _localTempOutputDir = NetworkPathHelper.CreateLocalTempOutputDirectory(_networkOutputDirectory);
            string localOutputPath = Path.Combine(_localTempOutputDir, fileName);

            try
            {
                // 让内部应用将文件生成到本地临时目录
                // 无论生成多少个文件（Excel 多 sheet），都在这个临时目录中
                _innerApplication.SaveAsPDF(localOutputPath);

                // 转换完成后，同步整个目录到网络路径
                NetworkPathHelper.SyncDirectoryToNetwork(_localTempOutputDir, _networkOutputDirectory);
            }
            finally
            {
                // 清理临时目录
                NetworkPathHelper.CleanupTempDirectory(_localTempOutputDir);
            }
        }
        else
        {
            // 本地路径输出：直接委托给内部应用
            _innerApplication.SaveAsPDF(toFilePath);
        }
    }

    public void CloseDocument()
    {
        _innerApplication.CloseDocument();
    }

    public void Dispose()
    {
        try
        {
            _innerApplication.Dispose();
        }
        finally
        {
            // 清理输入临时文件
            NetworkPathHelper.CleanupTempFile(_tempInputFilePath);

            // 清理输出临时目录（如果还存在的话）
            if (!string.IsNullOrEmpty(_localTempOutputDir))
            {
                NetworkPathHelper.CleanupTempDirectory(_localTempOutputDir);
            }
        }
    }
}

/// <summary>
/// 网络路径处理工具类，专门解决 Office COM 组件在 UNC 网络路径上的兼容性问题
/// 提供网络路径检测、临时文件管理、文件同步等功能
/// </summary>
public static class NetworkPathHelper
{
    /// <summary>
    /// 检查路径是否为 UNC 网络路径
    /// </summary>
    /// <param name="path">要检查的路径</param>
    /// <returns>如果是网络路径返回 true，否则返回 false</returns>
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
    /// <param name="networkFilePath">网络文件路径</param>
    /// <returns>本地临时文件路径</returns>
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
    /// <param name="networkOutputPath">网络输出文件路径</param>
    /// <returns>本地临时文件路径</returns>
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
    /// 为网络输出目录创建本地临时映射目录
    /// 用于支持多文件输出场景（如 Excel 多 sheet 导出）
    /// </summary>
    /// <param name="networkDirectory">网络目录路径</param>
    /// <returns>本地临时目录路径</returns>
    public static string CreateLocalTempOutputDirectory(string networkDirectory)
    {
        if (!IsNetworkPath(networkDirectory))
            return networkDirectory; // 如果不是网络路径，直接返回原路径

        try
        {
            // 创建基于目录路径哈希的唯一临时目录
            using (var md5 = MD5.Create())
            {
                byte[] hash = md5.ComputeHash(Encoding.UTF8.GetBytes(networkDirectory));
                string hashString = BitConverter.ToString(hash).Replace("-", "").Substring(0, 8);

                string tempDir = Path.Combine(Path.GetTempPath(), "Office2PDF_NetworkOutput", $"Dir_{hashString}_{Guid.NewGuid():N}");

                if (!Directory.Exists(tempDir))
                    Directory.CreateDirectory(tempDir);

                return tempDir;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"无法创建网络目录的本地临时映射: {networkDirectory}, 错误: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 将临时文件复制到网络路径
    /// </summary>
    /// <param name="localTempPath">本地临时文件路径</param>
    /// <param name="networkPath">目标网络文件路径</param>
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
    /// 将本地临时目录中的所有文件同步到网络目录
    /// 用于支持多文件输出场景（如 Excel 多 sheet 导出）
    /// </summary>
    /// <param name="localTempDir">本地临时目录</param>
    /// <param name="networkDir">目标网络目录</param>
    public static void SyncDirectoryToNetwork(string localTempDir, string networkDir)
    {
        try
        {
            // 确保网络目录存在
            if (!Directory.Exists(networkDir))
            {
                Directory.CreateDirectory(networkDir);
            }

            // 复制所有生成的文件到网络路径
            foreach (var file in Directory.GetFiles(localTempDir))
            {
                string fileName = Path.GetFileName(file);
                string targetPath = Path.Combine(networkDir, fileName);
                File.Copy(file, targetPath, true);
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"无法同步文件到网络目录: {networkDir}, 错误: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 清理临时文件
    /// </summary>
    /// <param name="tempFilePath">临时文件路径</param>
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
    /// 清理临时目录及其所有内容
    /// </summary>
    /// <param name="tempDir">临时目录路径</param>
    public static void CleanupTempDirectory(string tempDir)
    {
        try
        {
            if (!string.IsNullOrEmpty(tempDir) && Directory.Exists(tempDir))
            {
                Directory.Delete(tempDir, true);
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

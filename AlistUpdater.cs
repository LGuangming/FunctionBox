using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FunctionBox
{
    public class AlistUpdater
    {
        // TODO: 需替换为您 Alist 的直链地址
        // 例如 VersionUrl 指向的 .txt 文件内容只写版本号如: 1.0.0.1
        public static string VersionUrl = "https://alist.guangming.pro/d/OSS/CloudFlare%20R2/Tools/FunctionBox/version.txt";
        
        // InstallerZipUrl 指向您的安装包压缩包直链 (也就是 VS 发布出来的所有文件打包成的 .zip)
        public static string InstallerZipUrl = "https://alist.guangming.pro/d/OSS/CloudFlare%20R2/Tools/FunctionBox/FunctionBox.zip";

        public static async Task CheckAndUpdateAsync()
        {
            try
            {
                string onlineVersionStr;
                using (WebClient client = new WebClient())
                {
                    // 设置TLS，防止 HTTPS 请求失败
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                    onlineVersionStr = await client.DownloadStringTaskAsync(VersionUrl);
                }

                if (!Version.TryParse(onlineVersionStr.Trim(), out Version onlineVersion))
                {
                    MessageBox.Show("解析线上版本号失败，请检查配置或网络。", "更新提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                Version localVersion = Assembly.GetExecutingAssembly().GetName().Version;
                if (onlineVersion > localVersion)
                {
                    if (MessageBox.Show($"发现新版本: {onlineVersion}\n当前版本: {localVersion}\n是否下载更新？",
                        "更新提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    {
                        await DownloadAndRunInstallerAsync();
                    }
                }
                else
                {
                    MessageBox.Show($"当前已是最新版本 ({localVersion})。", "更新提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"检查更新失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static async Task DownloadAndRunInstallerAsync()
        {
            try
            {
                string tempDir = Path.Combine(Path.GetTempPath(), $"FunctionBoxUpdate_{DateTime.Now.Ticks}");
                Directory.CreateDirectory(tempDir);

                string saveZipPath = Path.Combine(tempDir, "update.zip");

                MessageBox.Show("点击确定后将开始下载并稍后自动启动。由于在后台下载，这可能需要一点时间，请稍候...", "下载中", MessageBoxButtons.OK, MessageBoxIcon.Information);

                using (WebClient client = new WebClient())
                {
                    await client.DownloadFileTaskAsync(new Uri(InstallerZipUrl), saveZipPath);
                }

                // 解压ZIP文件
                ZipFile.ExtractToDirectory(saveZipPath, tempDir);

                // 在解压目录寻找 setup.exe
                string setupPath = Path.Combine(tempDir, "setup.exe");
                
                if (File.Exists(setupPath))
                {
                    MessageBox.Show("下载完成！点击【确定】后将启动更新程序。\n\n⚠️ 重要提示: 请不要忘记在安装开始前关闭所有已打开的 Word 窗口以保证成功覆盖文件。", "更新准备就绪", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Process.Start(setupPath);
                }
                else
                {
                    MessageBox.Show("更新包解压成功，但未找到 setup.exe，请确认上传的 zip 中包含该文件。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"下载或执行更新失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}

using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;

namespace FunctionBox
{
    public static class AlistUpdater
    {
        private const string LatestReleaseApiUrl = "https://api.github.com/repos/LGuangming/FunctionBox/releases/latest";
        private const string PreferredAssetName = "FunctionBox.zip";

        public static async Task CheckAndUpdateAsync()
        {
            using (UpdateStatusForm statusForm = new UpdateStatusForm())
            {
                try
                {
                    statusForm.UpdateStatus("正在检查更新...");
                    statusForm.SetIndeterminate(true);
                    statusForm.Show();
                    statusForm.Activate();

                    ReleaseInfo release = await GetLatestReleaseAsync();
                    Version localVersion = GetLocalVersion();

                    if (release.Version <= localVersion)
                    {
                        statusForm.Close();
                        ShowTopMostMessage(
                            $"当前已是最新版本（{FormatVersion(localVersion)}）。",
                            "更新提示",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        return;
                    }

                    statusForm.Close();
                    DialogResult result = ShowTopMostMessage(
                        $"发现新版本：{release.TagName}\n当前版本：{FormatVersion(localVersion)}\n\n是否现在下载并更新？",
                        "更新提示",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information);

                    if (result != DialogResult.Yes)
                    {
                        return;
                    }

                    statusForm.UpdateStatus("正在下载更新包...");
                    statusForm.SetIndeterminate(false);
                    statusForm.UpdateProgress(0, "准备下载...");
                    statusForm.Show();
                    statusForm.Activate();

                    string setupPath = await DownloadAndPrepareInstallerAsync(release, statusForm);

                    statusForm.UpdateStatus("更新包准备完成，正在启动安装程序...");
                    statusForm.SetIndeterminate(true);
                    statusForm.Close();
                    ShowTopMostMessage(
                        "更新包准备完成，点击“确定”后将启动安装程序。\n\n开始安装前请关闭所有 Word 窗口，以便顺利覆盖旧文件。",
                        "更新准备完成",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                    Process.Start(setupPath);
                }
                catch (Exception ex)
                {
                    statusForm.Close();
                    ShowTopMostMessage(
                        $"检查或执行更新失败：{ex.Message}",
                        "更新失败",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private static async Task<ReleaseInfo> GetLatestReleaseAsync()
        {
            using (WebClient client = CreateWebClient())
            {
                string json = await client.DownloadStringTaskAsync(LatestReleaseApiUrl);
                JObject releaseObject = JObject.Parse(json);

                string tagName = releaseObject["tag_name"]?.ToString();
                if (!TryParseReleaseVersion(tagName, out Version version))
                {
                    throw new InvalidOperationException("无法解析 GitHub Release 版本号。");
                }

                JArray assets = releaseObject["assets"] as JArray;
                if (assets == null || assets.Count == 0)
                {
                    throw new InvalidOperationException("最新 Release 未包含可下载资源。");
                }

                JObject asset = assets
                    .OfType<JObject>()
                    .FirstOrDefault(item => string.Equals(item["name"]?.ToString(), PreferredAssetName, StringComparison.OrdinalIgnoreCase))
                    ?? assets.OfType<JObject>().FirstOrDefault(item => item["name"]?.ToString().EndsWith(".zip", StringComparison.OrdinalIgnoreCase) == true);

                if (asset == null)
                {
                    throw new InvalidOperationException("最新 Release 中未找到 zip 更新包。");
                }

                string assetUrl = asset["browser_download_url"]?.ToString();
                if (string.IsNullOrWhiteSpace(assetUrl))
                {
                    throw new InvalidOperationException("Release 资源下载地址无效。");
                }

                return new ReleaseInfo
                {
                    TagName = tagName,
                    Version = version,
                    AssetName = asset["name"]?.ToString(),
                    AssetUrl = assetUrl
                };
            }
        }

        private static async Task<string> DownloadAndPrepareInstallerAsync(ReleaseInfo release, UpdateStatusForm statusForm)
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "FunctionBoxUpdate_" + DateTime.Now.Ticks);
            Directory.CreateDirectory(tempDir);

            string zipPath = Path.Combine(tempDir, release.AssetName ?? PreferredAssetName);
            await DownloadFileWithProgressAsync(release.AssetUrl, zipPath, statusForm);

            statusForm.UpdateStatus("正在解压更新包...");
            statusForm.SetIndeterminate(true);
            string extractDir = Path.Combine(tempDir, "package");
            ZipFile.ExtractToDirectory(zipPath, extractDir);

            string setupPath = Directory
                .GetFiles(extractDir, "setup.exe", SearchOption.AllDirectories)
                .FirstOrDefault();

            if (string.IsNullOrWhiteSpace(setupPath))
            {
                throw new FileNotFoundException("更新包中未找到 setup.exe。");
            }

            return setupPath;
        }

        private static WebClient CreateWebClient()
        {
            ServicePointManager.SecurityProtocol =
                SecurityProtocolType.Tls12 |
                SecurityProtocolType.Tls11 |
                SecurityProtocolType.Tls;

            WebClient client = new WebClient();
            client.Headers[HttpRequestHeader.UserAgent] = "FunctionBox-Updater";
            client.Headers[HttpRequestHeader.Accept] = "application/vnd.github+json";
            return client;
        }
        private static Task DownloadFileWithProgressAsync(string url, string destinationPath, UpdateStatusForm statusForm)
        {
            TaskCompletionSource<bool> tcs = new TaskCompletionSource<bool>();
            WebClient client = CreateWebClient();

            client.DownloadProgressChanged += (sender, args) =>
            {
                statusForm.UpdateProgress(
                    args.ProgressPercentage,
                    $"正在下载更新包... {FormatSize(args.BytesReceived)} / {FormatSize(args.TotalBytesToReceive)}");
            };

            client.DownloadFileCompleted += (sender, args) =>
            {
                client.Dispose();

                if (args.Cancelled)
                {
                    tcs.TrySetCanceled();
                    return;
                }

                if (args.Error != null)
                {
                    tcs.TrySetException(args.Error);
                    return;
                }

                statusForm.UpdateProgress(100, "下载完成");
                tcs.TrySetResult(true);
            };

            client.DownloadFileAsync(new Uri(url), destinationPath);
            return tcs.Task;
        }

        private static Version GetLocalVersion()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            return assembly.GetName().Version ?? new Version(0, 0, 0, 0);
        }

        private static bool TryParseReleaseVersion(string tagName, out Version version)
        {
            version = new Version(0, 0, 0, 0);
            if (string.IsNullOrWhiteSpace(tagName))
            {
                return false;
            }

            string normalized = tagName.Trim();
            if (normalized.StartsWith("v", StringComparison.OrdinalIgnoreCase))
            {
                normalized = normalized.Substring(1);
            }

            string[] parts = normalized.Split('.');
            if (parts.Length < 2 || parts.Length > 4)
            {
                return false;
            }

            int[] numbers = new int[4];
            for (int index = 0; index < parts.Length; index++)
            {
                if (!int.TryParse(parts[index], out numbers[index]))
                {
                    return false;
                }
            }

            version = new Version(numbers[0], numbers[1], numbers[2], numbers[3]);
            return true;
        }

        private static string FormatVersion(Version version)
        {
            return $"v{version.Major}.{version.Minor}.{version.Build}";
        }
        private static string FormatSize(long bytes)
        {
            if (bytes < 0)
            {
                return "?";
            }

            string[] units = { "B", "KB", "MB", "GB" };
            double size = bytes;
            int unitIndex = 0;

            while (size >= 1024 && unitIndex < units.Length - 1)
            {
                size /= 1024;
                unitIndex++;
            }

            return $"{size:0.##} {units[unitIndex]}";
        }

        private static DialogResult ShowTopMostMessage(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            using (TopMostMessageOwner owner = new TopMostMessageOwner())
            {
                owner.Show();
                owner.Activate();
                return MessageBox.Show(owner, text, caption, buttons, icon);
            }
        }

        private sealed class ReleaseInfo
        {
            public string TagName { get; set; }
            public Version Version { get; set; }
            public string AssetName { get; set; }
            public string AssetUrl { get; set; }
        }

        private sealed class TopMostMessageOwner : Form
        {
            public TopMostMessageOwner()
            {
                ShowInTaskbar = false;
                StartPosition = FormStartPosition.CenterScreen;
                FormBorderStyle = FormBorderStyle.FixedToolWindow;
                Size = new System.Drawing.Size(1, 1);
                Opacity = 0;
                TopMost = true;
            }
        }

        private sealed class UpdateStatusForm : Form
        {
            private readonly Label statusLabel;
            private readonly ProgressBar progressBar;

            public UpdateStatusForm()
            {
                Text = "更新中";
                StartPosition = FormStartPosition.CenterScreen;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                ShowInTaskbar = false;
                TopMost = true;
                Width = 360;
                Height = 130;

                statusLabel = new Label
                {
                    AutoSize = false,
                    Left = 20,
                    Top = 18,
                    Width = 300,
                    Height = 24,
                    Text = "准备中..."
                };

                progressBar = new ProgressBar
                {
                    Left = 20,
                    Top = 55,
                    Width = 300,
                    Style = ProgressBarStyle.Marquee,
                    MarqueeAnimationSpeed = 30
                };

                Controls.Add(statusLabel);
                Controls.Add(progressBar);
            }

            public void UpdateStatus(string message)
            {
                statusLabel.Text = message;
                statusLabel.Refresh();
            }
            public void SetIndeterminate(bool isIndeterminate)
            {
                progressBar.Style = isIndeterminate ? ProgressBarStyle.Marquee : ProgressBarStyle.Continuous;
                progressBar.MarqueeAnimationSpeed = isIndeterminate ? 30 : 0;
                if (!isIndeterminate)
                {
                    progressBar.Value = 0;
                }
            }
            public void UpdateProgress(int percentage, string message)
            {
                statusLabel.Text = message;
                if (progressBar.Style != ProgressBarStyle.Continuous)
                {
                    progressBar.Style = ProgressBarStyle.Continuous;
                    progressBar.MarqueeAnimationSpeed = 0;
                }

                progressBar.Value = Math.Max(0, Math.Min(100, percentage));
                statusLabel.Refresh();
                progressBar.Refresh();
            }
        }
    }
}

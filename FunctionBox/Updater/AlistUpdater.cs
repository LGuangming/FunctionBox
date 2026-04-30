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
        private const string GitHubCdnPrefix = "https://gh.api.99988866.xyz/";
        private const string PreferredAssetName = "FunctionBox.zip";

        public static async Task CheckAndUpdateAsync()
        {
            UpdateStatusForm statusForm = null;
            try
            {
                statusForm = new UpdateStatusForm();
                Version localVersion = GetLocalVersion();

                statusForm.SetVersionText($"当前版本：{FormatVersion(localVersion)}");
                statusForm.UpdateStatus("正在检查更新...");
                statusForm.SetIndeterminate(true);
                statusForm.SetActionButtonState(false, "正在检查...");
                statusForm.SafeShow();
                statusForm.SafeActivate();

                ReleaseInfo release;
                try
                {
                    release = await GetLatestReleaseAsync();
                }
                catch (Exception ex)
                {
                    statusForm.UpdateStatus($"检查失败：{ex.Message}");
                    statusForm.SetIndeterminate(false);
                    statusForm.SetActionButtonState(true, "关闭", isClose: true);
                    await statusForm.WaitForUserActionAsync();
                    return;
                }

                if (release.Version <= localVersion)
                {
                    statusForm.UpdateStatus($"当前已是最新版本");
                    statusForm.SetIndeterminate(false);
                    statusForm.SetActionButtonState(true, "关闭", isClose: true);
                    await statusForm.WaitForUserActionAsync();
                    return;
                }

                statusForm.SafeHide();
                DialogResult confirmResult = ShowTopMostMessage(
                    $"发现新版本：{release.TagName}\n当前版本：{FormatVersion(localVersion)}\n\n是否现在下载并更新？",
                    "更新提示",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);
                if (confirmResult != DialogResult.Yes)
                {
                    return;
                }

                statusForm.SetVersionText($"当前版本：{FormatVersion(localVersion)}    最新版本：{release.TagName}");
                statusForm.UpdateStatus("正在下载更新包...");
                statusForm.SetIndeterminate(false);
                statusForm.SetActionButtonState(false, "下载中...");
                statusForm.UpdateProgress(0, "准备下载...");
                statusForm.SafeShow();
                statusForm.SafeActivate();

                bool proceed = true;
                if (!proceed)
                {
                    return;
                }

                string setupPath = await DownloadAndPrepareInstallerAsync(release, statusForm);

                statusForm.UpdateStatus("就绪！开始安装前请关闭所有 Word 窗口。");
                statusForm.SetIndeterminate(true);
                statusForm.SetActionButtonState(true, "开始安装", isClose: true);

                await statusForm.WaitForUserActionAsync();

                Process.Start(setupPath);
            }
            catch (Exception ex)
            {
                if (statusForm != null && !statusForm.IsDisposed)
                {
                    statusForm.UpdateStatus($"错误：{ex.Message}");
                    statusForm.SetIndeterminate(false);
                    statusForm.SetActionButtonState(true, "关闭", isClose: true);
                    await statusForm.WaitForUserActionAsync();
                }
            }
            finally
            {
                if (statusForm != null && !statusForm.IsDisposed)
                {
                    if (statusForm.InvokeRequired)
                    {
                        statusForm.Invoke(new Action(statusForm.Dispose));
                    }
                    else
                    {
                        statusForm.Dispose();
                    }
                }
            }
        }

        private static async Task<ReleaseInfo> GetLatestReleaseAsync()
        {
            Exception lastException = null;

            foreach (string apiUrl in GetCandidateUrls(LatestReleaseApiUrl))
            {
                try
                {
                    using (WebClient client = CreateWebClient())
                    {
                        string json = await client.DownloadStringTaskAsync(apiUrl);
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
                catch (Exception ex)
                {
                    lastException = ex;
                }
            }

            throw new InvalidOperationException(
                "无法获取最新 Release 信息，请检查 GitHub 或 CDN 网络连接。",
                lastException);
        }

        private static async Task<string> DownloadAndPrepareInstallerAsync(ReleaseInfo release, UpdateStatusForm statusForm)
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "FunctionBoxUpdate_" + DateTime.Now.Ticks);
            Directory.CreateDirectory(tempDir);

            string zipPath = Path.Combine(tempDir, release.AssetName ?? PreferredAssetName);
            await DownloadFileWithProgressAsync(GetCandidateUrls(release.AssetUrl), zipPath, statusForm);

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
        private static async Task DownloadFileWithProgressAsync(string[] urls, string destinationPath, UpdateStatusForm statusForm)
        {
            Exception lastException = null;

            foreach (string url in urls)
            {
                try
                {
                    await DownloadFileWithProgressCoreAsync(url, destinationPath, statusForm);
                    return;
                }
                catch (Exception ex)
                {
                    lastException = ex;
                }
            }

            throw new InvalidOperationException("无法下载更新包，请检查 GitHub 或 CDN 网络连接。", lastException);
        }
        private static Task DownloadFileWithProgressCoreAsync(string url, string destinationPath, UpdateStatusForm statusForm)
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
                if (args.Cancelled)
                {
                    client.Dispose();
                    tcs.TrySetCanceled();
                    return;
                }

                if (args.Error != null)
                {
                    client.Dispose();
                    tcs.TrySetException(args.Error);
                    return;
                }

                client.Dispose();
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
        private static string[] GetCandidateUrls(string originalUrl)
        {
            if (string.IsNullOrWhiteSpace(originalUrl))
            {
                return Array.Empty<string>();
            }

            return new[]
            {
                originalUrl,
                BuildCdnUrl(originalUrl)
            }
            .Where(url => !string.IsNullOrWhiteSpace(url))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        }
        private static string BuildCdnUrl(string originalUrl)
        {
            if (string.IsNullOrWhiteSpace(originalUrl))
            {
                return null;
            }

            if (originalUrl.StartsWith(GitHubCdnPrefix, StringComparison.OrdinalIgnoreCase))
            {
                return originalUrl;
            }

            return GitHubCdnPrefix + originalUrl;
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
            private readonly Label versionLabel;
            private readonly ProgressBar progressBar;
            private readonly Button actionButton;
            private TaskCompletionSource<bool> _tcs;
            private bool _isCloseAction;

            public UpdateStatusForm()
            {
                Text = "FunctionBox 更新程序";
                StartPosition = FormStartPosition.CenterScreen;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                ShowInTaskbar = false;
                TopMost = true;
                ClientSize = new System.Drawing.Size(340, 180);

                statusLabel = new Label
                {
                    AutoSize = false,
                    Left = 20,
                    Top = 16,
                    Width = 300,
                    Height = 42,
                    Text = "准备中...",
                    TextAlign = System.Drawing.ContentAlignment.TopLeft
                };

                versionLabel = new Label
                {
                    AutoSize = false,
                    Left = 20,
                    Top = 60,
                    Width = 300,
                    Height = 20,
                    Text = "当前版本：未知"
                };

                progressBar = new ProgressBar
                {
                    Left = 20,
                    Top = 85,
                    Width = 300,
                    Height = 15,
                    Style = ProgressBarStyle.Marquee,
                    MarqueeAnimationSpeed = 30
                };

                actionButton = new Button
                {
                    Width = 100,
                    Height = 30,
                    Text = "...",
                    FlatStyle = FlatStyle.System,
                    Enabled = false
                };

                // 动态计算位置以彻底解决缩放带来的偏移问题
                this.Load += (s, e) => AdjustButtonPosition();
                this.Resize += (s, e) => AdjustButtonPosition();
                actionButton.Click += (s, e) =>
                {
                    if (_isCloseAction)
                    {
                        this.Close();
                    }
                    else
                    {
                        _tcs?.TrySetResult(true);
                    }
                };

                Controls.Add(statusLabel);
                Controls.Add(versionLabel);
                Controls.Add(progressBar);
                Controls.Add(actionButton);

                this.FormClosed += (s, e) =>
                {
                    _tcs?.TrySetResult(false);
                };
            }

            private void AdjustButtonPosition()
            {
                if (actionButton == null) return;
                // 绝对水平居中
                actionButton.Left = (this.ClientSize.Width - actionButton.Width) / 2;
                // 固定距离底部 15 像素
                actionButton.Top = this.ClientSize.Height - actionButton.Height - 15;
            }

            public Task<bool> WaitForUserActionAsync()
            {
                _tcs = new TaskCompletionSource<bool>();
                return _tcs.Task;
            }

            public void SetActionButtonState(bool enabled, string text, bool isClose = false)
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action<bool, string, bool>(SetActionButtonState), enabled, text, isClose);
                    return;
                }
                actionButton.Enabled = enabled;
                actionButton.Text = text;
                _isCloseAction = isClose;
            }

            public void UpdateStatus(string message)
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action<string>(UpdateStatus), message);
                    return;
                }
                statusLabel.Text = message;
                statusLabel.Refresh();
            }
            public void SetVersionText(string message)
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action<string>(SetVersionText), message);
                    return;
                }
                versionLabel.Text = message;
                versionLabel.Refresh();
            }
            public void SetIndeterminate(bool isIndeterminate)
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action<bool>(SetIndeterminate), isIndeterminate);
                    return;
                }
                progressBar.Style = isIndeterminate ? ProgressBarStyle.Marquee : ProgressBarStyle.Continuous;
                progressBar.MarqueeAnimationSpeed = isIndeterminate ? 30 : 0;
                if (!isIndeterminate)
                {
                    progressBar.Value = 0;
                }
            }
            public void UpdateProgress(int percentage, string message)
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action<int, string>(UpdateProgress), percentage, message);
                    return;
                }
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

            public void SafeShow()
            {
                if (this.InvokeRequired) { this.Invoke(new Action(SafeShow)); return; }
                this.Show();
            }

            public void SafeHide()
            {
                if (this.InvokeRequired) { this.Invoke(new Action(SafeHide)); return; }
                this.Hide();
            }

            public void SafeActivate()
            {
                if (this.InvokeRequired) { this.Invoke(new Action(SafeActivate)); return; }
                this.Activate();
            }
        }
    }
}

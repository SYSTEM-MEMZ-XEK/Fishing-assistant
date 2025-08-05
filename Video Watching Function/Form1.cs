using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Newtonsoft.Json;
using Vlc.DotNet.Core;
using Vlc.DotNet.Forms;

namespace VideoPlayer
{
    public class VideoPlayerForm : Form
    {
        private List<string> videoFiles = new List<string>();
        private int currentVideoIndex = 0;
        private FishingSettings settings = new FishingSettings();
        private VlcControl vlcControl;
        private bool isFullScreen = false;
        private Point originalLocation;
        private Size originalSize;
        private FormBorderStyle originalBorderStyle;
        private Timer volumeDisplayTimer;
        private Label volumeLabel;
        private Label positionLabel;
        private string libVlcPath;

        public VideoPlayerForm(string videoPath)
        {
            if (!InitializeVlcPlayer())
            {
                this.Close();
                return;
            }

            LoadSettings();
            LoadVideos(videoPath);
            ApplySettings();
            PlayCurrentVideo();
        }

        private bool InitializeVlcPlayer()
        {
            try
            {
                // 设置VLC库路径
                libVlcPath = GetLibVlcDirectory();
                if (libVlcPath == null)
                {
                    return false;
                }

                // 初始化VLC控件
                var options = new string[]
                {
                    ":avcodec-hw=dxva2",  // 使用DXVA2硬件加速
                    ":network-caching=1000",  // 增加网络缓存
                    ":file-caching=1000",     // 增加文件缓存
                    ":drop-late-frames",      // 丢弃延迟帧
                    ":skip-frames"            // 允许跳帧
                };

                vlcControl = new VlcControl
                {
                    VlcLibDirectory = new DirectoryInfo(libVlcPath),
                    VlcMediaplayerOptions = options
                };

                vlcControl.BeginInit();
                vlcControl.Dock = DockStyle.Fill;
                this.Controls.Add(vlcControl);
                vlcControl.EndInit();

                // 添加音量显示标签
                volumeLabel = new Label
                {
                    AutoSize = true,
                    BackColor = Color.FromArgb(100, 0, 0, 0),
                    ForeColor = Color.White,
                    Font = new Font("Arial", 12, FontStyle.Bold),
                    Visible = false
                };
                this.Controls.Add(volumeLabel);
                volumeLabel.BringToFront();

                // 添加进度显示标签
                positionLabel = new Label
                {
                    AutoSize = true,
                    BackColor = Color.FromArgb(100, 0, 0, 0),
                    ForeColor = Color.White,
                    Font = new Font("Arial", 10),
                    Visible = false,
                    Location = new Point(10, 40)
                };
                this.Controls.Add(positionLabel);
                positionLabel.BringToFront();

                // 初始化音量显示定时器
                volumeDisplayTimer = new Timer { Interval = 2000 };
                volumeDisplayTimer.Tick += (s, e) =>
                {
                    volumeLabel.Visible = false;
                    positionLabel.Visible = false;
                    volumeDisplayTimer.Stop();
                };

                // 窗体设置
                this.Text = "视频播放器";
                this.Size = new Size(800, 600);
                this.FormBorderStyle = FormBorderStyle.Sizable;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.KeyPreview = true;
                this.KeyDown += MainForm_KeyDown;

                // 事件处理
                vlcControl.PositionChanged += (s, e) => UpdatePositionDisplay();
                vlcControl.Stopped += (s, e) => {
                    if (settings.VideoAutoSwitch)
                    {
                        PlayNextVideo();
                    }
                };

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化VLC播放器失败: {ex.Message}\n\n" +
                                "请确保VLC库文件已正确放置。",
                                "初始化错误",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                return false;
            }
        }

        private void UpdatePositionDisplay()
        {
            if (vlcControl == null || !vlcControl.IsPlaying) return;

            TimeSpan current = TimeSpan.FromMilliseconds(vlcControl.Time);
            TimeSpan total = TimeSpan.FromMilliseconds(vlcControl.Length);
            string positionText = $"{current:mm\\:ss} / {total:mm\\:ss}";
            Point position = new Point(10, this.ClientSize.Height - 40);

            // 使用线程安全方式更新UI
            void SafeUpdate()
            {
                positionLabel.Text = positionText;
                positionLabel.Location = position;
                positionLabel.Visible = true;
            }

            // 检查是否需要跨线程调用
            if (positionLabel.InvokeRequired)
            {
                positionLabel.Invoke((MethodInvoker)SafeUpdate);
            }
            else
            {
                SafeUpdate();
            }

            // 重置定时器
            volumeDisplayTimer.Stop();
            volumeDisplayTimer.Start();
        }

        private void ShowVolumeDisplay()
        {
            if (vlcControl == null) return;

            string volumeText = $"音量: {vlcControl.Audio.Volume}%";
            Point position = new Point(10, 10);

            void SafeUpdate()
            {
                volumeLabel.Text = volumeText;
                volumeLabel.Location = position;
                volumeLabel.Visible = true;
            }

            if (volumeLabel.InvokeRequired)
            {
                volumeLabel.Invoke((MethodInvoker)SafeUpdate);
            }
            else
            {
                SafeUpdate();
            }

            volumeDisplayTimer.Stop();
            volumeDisplayTimer.Start();
        }

        private string GetLibVlcDirectory()
        {
            string baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string libDir = Path.Combine(baseDir, "libvlc");

            // 检查是否存在libvlc目录
            if (!Directory.Exists(libDir))
            {
                MessageBox.Show($"未找到VLC库文件目录: {libDir}\n\n" +
                                "请确保libvlc目录与应用程序位于同一目录下。",
                                "文件缺失错误",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                return null;
            }

            // 根据系统架构选择子目录
            string platformDir = Environment.Is64BitProcess ? "win-x64" : "win-x86";
            string fullPath = Path.Combine(libDir, platformDir);

            if (!Directory.Exists(fullPath))
            {
                MessageBox.Show($"未找到平台特定VLC库: {fullPath}\n\n" +
                                $"请确保 {platformDir} 目录存在于libvlc中。",
                                "文件缺失错误",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                return null;
            }

            return fullPath;
        }

        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            // 空格键暂停/播放
            if (e.KeyCode == Keys.Space)
            {
                TogglePlayPause();
                e.Handled = true;
            }
            // 右箭头下一个视频
            else if (e.KeyCode == Keys.Right)
            {
                PlayNextVideo();
                e.Handled = true;
            }
            // 左箭头上一个视频
            else if (e.KeyCode == Keys.Left)
            {
                PlayPreviousVideo();
                e.Handled = true;
            }
            // ESC键退出全屏
            else if (e.KeyCode == Keys.Escape && isFullScreen)
            {
                ToggleFullScreen();
                e.Handled = true;
            }
            // F键切换全屏
            else if (e.KeyCode == Keys.F)
            {
                ToggleFullScreen();
                e.Handled = true;
            }
            // M键静音
            else if (e.KeyCode == Keys.M)
            {
                ToggleMute();
                e.Handled = true;
            }
            // 上箭头增加音量
            else if (e.KeyCode == Keys.Up)
            {
                IncreaseVolume();
                e.Handled = true;
            }
            // 下箭头减少音量
            else if (e.KeyCode == Keys.Down)
            {
                DecreaseVolume();
                e.Handled = true;
            }
            // 快进10秒
            else if (e.KeyCode == Keys.PageDown)
            {
                SeekForward(10000);
                e.Handled = true;
            }
            // 后退10秒
            else if (e.KeyCode == Keys.PageUp)
            {
                SeekBackward(10000);
                e.Handled = true;
            }
            // 右Ctrl + → 快进30秒
            else if (e.KeyCode == Keys.Right && e.Control)
            {
                SeekForward(30000);
                e.Handled = true;
            }
            // 右Ctrl + ← 后退30秒
            else if (e.KeyCode == Keys.Left && e.Control)
            {
                SeekBackward(30000);
                e.Handled = true;
            }
            // Q键退出
            else if (e.KeyCode == Keys.Q)
            {
                this.Close();
            }
        }

        private void TogglePlayPause()
        {
            if (vlcControl == null) return;

            if (vlcControl.IsPlaying)
            {
                vlcControl.Pause();
            }
            else
            {
                vlcControl.Play();
            }
        }

        private void PlayNextVideo()
        {
            if (videoFiles.Count == 0) return;

            currentVideoIndex = (currentVideoIndex + 1) % videoFiles.Count;
            PlayCurrentVideo();
        }

        private void PlayPreviousVideo()
        {
            if (videoFiles.Count == 0) return;

            currentVideoIndex = (currentVideoIndex - 1 + videoFiles.Count) % videoFiles.Count;
            PlayCurrentVideo();
        }

        private void ToggleFullScreen()
        {
            if (isFullScreen)
            {
                // 退出全屏
                this.FormBorderStyle = originalBorderStyle;
                this.Location = originalLocation;
                this.Size = originalSize;
                this.TopMost = false;
                isFullScreen = false;

                // 恢复控件位置
                volumeLabel.Location = new Point(10, 10);
                positionLabel.Location = new Point(10, this.ClientSize.Height - 40);
            }
            else
            {
                // 进入全屏
                originalBorderStyle = this.FormBorderStyle;
                originalLocation = this.Location;
                originalSize = this.Size;

                this.FormBorderStyle = FormBorderStyle.None;
                this.WindowState = FormWindowState.Maximized;
                this.TopMost = true;
                isFullScreen = true;

                // 调整控件位置
                volumeLabel.Location = new Point(20, 20);
                positionLabel.Location = new Point(20, this.ClientSize.Height - 40);
            }
        }

        private void ToggleMute()
        {
            if (vlcControl == null) return;

            vlcControl.Audio.IsMute = !vlcControl.Audio.IsMute;

            volumeLabel.Text = vlcControl.Audio.IsMute ? "静音" : "取消静音";
            volumeLabel.Location = new Point(10, 10);
            volumeLabel.Visible = true;

            volumeDisplayTimer.Stop();
            volumeDisplayTimer.Start();
        }

        private void IncreaseVolume()
        {
            if (vlcControl == null) return;

            int newVolume = Math.Min(vlcControl.Audio.Volume + 10, 200);
            vlcControl.Audio.Volume = newVolume;
            ShowVolumeDisplay();
        }

        private void DecreaseVolume()
        {
            if (vlcControl == null) return;

            int newVolume = Math.Max(vlcControl.Audio.Volume - 10, 0);
            vlcControl.Audio.Volume = newVolume;
            ShowVolumeDisplay();
        }

        private void SeekForward(long milliseconds)
        {
            if (vlcControl == null || !vlcControl.IsPlaying) return;

            long newTime = Math.Min(vlcControl.Time + milliseconds, vlcControl.Length);
            vlcControl.Time = newTime;
            UpdatePositionDisplay();
        }

        private void SeekBackward(long milliseconds)
        {
            if (vlcControl == null || !vlcControl.IsPlaying) return;

            long newTime = Math.Max(vlcControl.Time - milliseconds, 0);
            vlcControl.Time = newTime;
            UpdatePositionDisplay();
        }

        private void LoadSettings()
        {
            string configPath = Path.Combine(Application.StartupPath, "..", "Fishing settings.json");
            configPath = Path.GetFullPath(configPath); // 获取绝对路径
            if (File.Exists(configPath))
            {
                try
                {
                    string json = File.ReadAllText(configPath);
                    settings = JsonConvert.DeserializeObject<FishingSettings>(json);
                }
                catch
                {
                    // 使用默认设置
                    settings = new FishingSettings();
                }
            }
            else
            {
                // 如果配置文件不存在，使用默认设置
                settings = new FishingSettings();
            }
        }

        private void LoadVideos(string videoPath)
        {
            if (string.IsNullOrEmpty(videoPath))
            {
                videoPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyVideos));
            }

            if (!Directory.Exists(videoPath))
            {
                MessageBox.Show($"视频目录不存在: {videoPath}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 支持的视频格式
            string[] videoExtensions = {
                ".mp4", ".avi", ".mkv", ".wmv", ".mov", ".flv",
                ".mpg", ".mpeg", ".m4v", ".webm", ".ts", ".m2ts"
            };

            videoFiles = Directory.GetFiles(videoPath, "*.*")
                .Where(f => videoExtensions.Contains(Path.GetExtension(f).ToLower()))
                .OrderBy(f => f)
                .ToList();

            if (videoFiles.Count == 0)
            {
                MessageBox.Show($"未找到视频文件: {videoPath}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void PlayCurrentVideo()
        {
            if (videoFiles.Count == 0 || vlcControl == null) return;

            try
            {
                this.Cursor = Cursors.WaitCursor;
                vlcControl.Stop();

                // 添加性能优化选项
                var mediaOptions = new string[]
                {
                    ":no-avformat-timestamp-hack",
                    ":avcodec-skiploopfilter=4",
                    ":avcodec-skipt=4",
                    ":avcodec-fast"
                };

                vlcControl.SetMedia(new FileInfo(videoFiles[currentVideoIndex]), mediaOptions);
                vlcControl.Play();
                this.Text = $"视频播放器 - {Path.GetFileName(videoFiles[currentVideoIndex])}";

                // 应用音量设置
                if (settings.InitialVolume >= 0)
                {
                    vlcControl.Audio.Volume = settings.InitialVolume;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"播放视频失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                PlayNextVideo();
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void ApplySettings()
        {
            // 应用窗口大小
            try
            {
                if (!string.IsNullOrEmpty(settings.VideoWindowSize))
                {
                    // 使用正确的分隔符（可能是'*'或'x'）
                    char[] separators = { '*', 'x', 'X', '×' };
                    string[] sizeParts = settings.VideoWindowSize.Split(separators);

                    if (sizeParts.Length >= 2)
                    {
                        int width = int.Parse(sizeParts[0].Trim());
                        int height = int.Parse(sizeParts[1].Trim());
                        this.Size = new Size(width, height);
                    }
                }
            }
            catch
            {
                // 使用默认大小
                this.Size = new Size(800, 600);
            }

            // 设置全屏模式
            if (settings.VideoFullScreen)
            {
                ToggleFullScreen();
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (vlcControl != null)
            {
                vlcControl.Stop();
                vlcControl.Dispose();
            }

            volumeDisplayTimer?.Stop();

            base.OnFormClosing(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                vlcControl?.Dispose();
                volumeDisplayTimer?.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    // 设置类
    public class FishingSettings
    {
        // 窗口大小
        public string VideoWindowSize { get; set; } = "800,600";

        // 自动切换
        public bool VideoAutoSwitch { get; set; } = false;

        // 初始音量
        public int InitialVolume { get; set; } = 80;

        // 新增：全屏模式
        public bool VideoFullScreen { get; set; } = false;

        // 新增：摸鱼时打开视频选项
        public bool OpenVideoWhenFishing { get; set; } = false;

        // 新增：视频文件夹路径
        public string VideoPath { get; set; } = "";
    }
}
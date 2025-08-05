using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static Fishing_assistant.Form1;

namespace Fishing_assistant
{
    public partial class Form1 : Form
    {
        private const string ConfigFile = "Fishing settings.json";
        private FishingSettings settings = new FishingSettings();
        private bool bossStatus = false; // false表示老板不在，true表示老板来了
        private List<IntPtr> minimizedWindows = new List<IntPtr>(); // 存储最小化的窗口句柄
        private bool wasMuted = false; // 记录静音前的状态
        private bool isInFishingMode = false; // 标记是否处于摸鱼模式
        private bool isProcessingHotkey = false; // 防止热键重复处理
        private DateTime fishingStartTime; // 记录摸鱼开始时间
        private DateTime lastBossVisitTime; // 记录上次老板来的时间
        private Timer fishingTimer; // 用于计算摸鱼时长

        // 热键注册API
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        // 窗口操作API
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        // 音量控制API
        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);

        // 获取系统音量API
        [DllImport("winmm.dll")]
        private static extern int waveOutGetVolume(IntPtr hwo, out uint dwVolume);

        // 检查窗口是否存在API
        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        // 常量定义
        private const int SW_MINIMIZE = 6;
        private const int SW_RESTORE = 9;
        private const int VK_VOLUME_MUTE = 0xAD;
        private const int KEYEVENTF_EXTENDEDKEY = 0x1;
        private const int KEYEVENTF_KEYUP = 0x2;
        private const int MOD_CONTROL = 0x0002;
        private const int MOD_ALT = 0x0001;
        private const int MOD_SHIFT = 0x0004;
        private const int WM_HOTKEY = 0x0312;

        // 系统托盘图标
        private NotifyIcon trayIcon;

        public Form1()
        {
            InitializeComponent();
            InitializeTrayIcon();
            LoadSettings();
            InitializeUI();
            RegisterBossKey();
            // 初始化计时器
            fishingTimer = new Timer();
            fishingTimer.Interval = 1000; // 1秒
            fishingTimer.Tick += FishingTimer_Tick;
        }

        private void FishingTimer_Tick(object sender, EventArgs e)
        {
            // 更新摸鱼时长
            if (isInFishingMode)
            {
                settings.FishingDuration = DateTime.Now - fishingStartTime;
                UpdateStatisticsDisplay();
            }
        }

        private void UpdateStatisticsDisplay()
        {
            // 更新统计显示
            string durationText = $"{settings.FishingDuration.Hours}时{settings.FishingDuration.Minutes}分{settings.FishingDuration.Seconds}秒";
            label24.Text = $"老板今天来了{settings.BossVisitCount}次，今天的摸鱼时长：{durationText}";
        }
        private void InitializeTrayIcon()
        {
            trayIcon = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                Text = "摸鱼助手",
                Visible = false
            };

            trayIcon.DoubleClick += (s, e) => {
                this.Show();
                this.WindowState = FormWindowState.Normal;
                this.BringToFront();
                isInFishingMode = false;
                trayIcon.Visible = false;
            };

            ContextMenuStrip contextMenu = new ContextMenuStrip();
            contextMenu.Items.Add("显示窗口", null, (s, e) => {
                this.Show();
                this.WindowState = FormWindowState.Normal;
                this.BringToFront();
                isInFishingMode = false;
                trayIcon.Visible = false;
            });
            contextMenu.Items.Add("退出", null, (s, e) => Application.Exit());

            trayIcon.ContextMenuStrip = contextMenu;
        }

        private void LoadSettings()
        {
            if (File.Exists(ConfigFile))
            {
                try
                {
                    string json = File.ReadAllText(ConfigFile);
                    settings = JsonConvert.DeserializeObject<FishingSettings>(json);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"加载配置失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void SaveSettings()
        {
            try
            {
                string json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                File.WriteAllText(ConfigFile, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存配置失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitializeUI()
        {
            // 确保老板键设置不为null
            if (settings.BossKey == null)
            {
                settings.BossKey = new HotkeySetting
                {
                    Ctrl = true,
                    Alt = true,
                    Key = Keys.F12
                };
            }

            // 初始化老板键显示
            label1.Text = "当前老板键：" + FormatHotkey(settings.BossKey);

            // 初始化路径显示
            textBox1.Text = settings.NovelPath;
            textBox2.Text = settings.VideoPath;

            // 初始化老板来时操作
            if (settings.OnBossComing.OpenApps.Count > 0) textBox3.Text = settings.OnBossComing.OpenApps[0];
            if (settings.OnBossComing.OpenFiles.Count > 0) textBox4.Text = settings.OnBossComing.OpenFiles[0];
            if (settings.OnBossComing.OpenWebs.Count > 0) textBox5.Text = settings.OnBossComing.OpenWebs[0];
            if (settings.OnBossComing.CloseApps.Count > 0) textBox8.Text = settings.OnBossComing.CloseApps[0];
            if (settings.OnBossComing.MinimizeApps.Count > 0) textBox6.Text = settings.OnBossComing.MinimizeApps[0];
            checkBox2.Checked = settings.OnBossComing.CloseNovel;
            checkBox3.Checked = settings.OnBossComing.CloseVideo;
            checkBox1.Checked = settings.OnBossComing.MuteWhenMinimize;

            // 初始化老板走后操作
            if (settings.OnBossLeaving.OpenApps.Count > 0) textBox12.Text = settings.OnBossLeaving.OpenApps[0];
            if (settings.OnBossLeaving.OpenFiles.Count > 0) textBox11.Text = settings.OnBossLeaving.OpenFiles[0];
            if (settings.OnBossLeaving.OpenWebs.Count > 0) textBox10.Text = settings.OnBossLeaving.OpenWebs[0];
            if (settings.OnBossLeaving.CloseApps.Count > 0) textBox9.Text = settings.OnBossLeaving.CloseApps[0];
            checkBox4.Checked = settings.OnBossLeaving.RestoreNovel;
            checkBox5.Checked = settings.OnBossLeaving.RestoreVideo;
            checkBox6.Checked = settings.OnBossLeaving.RestoreClosedApps;

            // 初始化小说设置
            label11.Text = "小说文件夹路径：" + settings.NovelPath;
            label18.Text = $"当前文本字体: {settings.NovelFont.Name}, {settings.NovelFont.Size}pt";
            label21.Text = $"当前文本颜色: {settings.NovelTextColor.Name}";
            checkBox8.Checked = settings.NovelWindowTransparent;
            checkBox9.Checked = settings.NovelAutoTurnPage;
            numericUpDown1.Value = settings.NovelTurnPageSeconds;
            textBox7.Text = $"{settings.NovelWindowSize.Width}*{settings.NovelWindowSize.Height}";

            // 初始化视频设置
            label25.Text = "视频文件夹路径：" + settings.VideoPath;
            checkBox10.Checked = settings.VideoAutoSwitch;
            textBox13.Text = $"{settings.VideoWindowSize.Width}*{settings.VideoWindowSize.Height}";

            // 初始化新控件
            checkBox11.Checked = settings.OpenNovelWhenFishing;
            checkBox12.Checked = settings.OpenVideoWhenFishing;

            // 初始化统计显示
            UpdateStatisticsDisplay();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            settings.NovelPath = textBox1.Text;
            SaveSettings();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            settings.VideoPath = textBox2.Text;
            SaveSettings();
        }
        private void RegisterBossKey()
        {
            UnregisterBossKey();

            if (settings.BossKey != null && settings.BossKey.Key != Keys.None)
            {
                try
                {
                    // 计算修饰键组合
                    int modifiers = 0;
                    if (settings.BossKey.Ctrl) modifiers |= MOD_CONTROL;
                    if (settings.BossKey.Alt) modifiers |= MOD_ALT;
                    if (settings.BossKey.Shift) modifiers |= MOD_SHIFT;

                    // 检查键值是否有效
                    if (settings.BossKey.Key != Keys.None)
                    {
                        if (!RegisterHotKey(this.Handle, 1, modifiers, (int)settings.BossKey.Key))
                        {
                            MessageBox.Show("老板键注册失败，可能已被其他程序占用。请尝试更换老板键。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"注册老板键失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                // 初始化默认老板键
                settings.BossKey = new HotkeySetting
                {
                    Ctrl = true,
                    Alt = true,
                    Key = Keys.F12
                };
                SaveSettings();
                RegisterBossKey();
            }
        }

        // 摸鱼时打开小说
        private void CheckBox11_CheckedChanged(object sender, EventArgs e)
        {
            settings.OpenNovelWhenFishing = checkBox11.Checked;
            SaveSettings();
        }

        // 摸鱼时打开视频
        private void CheckBox12_CheckedChanged(object sender, EventArgs e)
        {
            settings.OpenVideoWhenFishing = checkBox12.Checked;
            SaveSettings();
        }
        private void UnregisterBossKey()
        {
            UnregisterHotKey(this.Handle, 1);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == 1)
            {
                // 防止重复处理
                if (isProcessingHotkey) return;
                isProcessingHotkey = true;

                try
                {
                    // 只有在摸鱼模式下才响应老板键
                    if (isInFishingMode)
                    {
                        // 老板键被按下
                        if (bossStatus)
                        {
                            // 当前状态是老板来了，那么老板走了
                            BossLeavingActions();
                        }
                        else
                        {
                            // 老板来了
                            BossComingActions();
                        }
                        bossStatus = !bossStatus;
                    }
                }
                finally
                {
                    isProcessingHotkey = false;
                }
            }
            base.WndProc(ref m);
        }

        private void BossComingActions()
        {
            // 执行老板来时的操作
            // 1. 打开设置的应用/文件/网页
            if (!string.IsNullOrEmpty(textBox3.Text))
                Process.Start(textBox3.Text);
            if (!string.IsNullOrEmpty(textBox4.Text))
                Process.Start(textBox4.Text);
            if (!string.IsNullOrEmpty(textBox5.Text))
                Process.Start("explorer.exe", textBox5.Text);

            // 2. 关闭设置的应用
            if (!string.IsNullOrEmpty(textBox8.Text))
            {
                string appName = Path.GetFileNameWithoutExtension(textBox8.Text);
                CloseApplication(appName);
            }

            // 3. 关闭小说和视频
            if (checkBox2.Checked) CloseApplication("Novel reading function");
            if (checkBox3.Checked) CloseApplication("Video Watching function");

            // 添加额外的关闭尝试
            if (checkBox2.Checked) CloseApplication("Novel reading function.exe");
            if (checkBox3.Checked) CloseApplication("Video Watching function.exe");

            // 4. 最小化应用
            if (!string.IsNullOrEmpty(textBox6.Text))
            {
                string appName = Path.GetFileNameWithoutExtension(textBox6.Text);
                MinimizeApplication(appName);
            }

            // 5. 静音
            if (checkBox1.Checked)
            {
                wasMuted = IsSystemMuted();
                MuteSystem(true);
            }

            // 增加老板来访计数
            settings.BossVisitCount++;
            SaveSettings();

            // 更新统计显示
            UpdateStatisticsDisplay();

            // 隐藏主窗口
            this.Hide();
        }

        private void BossLeavingActions()
        {
            // 执行老板走后的操作
            // 1. 打开设置的应用/文件/网页
            if (!string.IsNullOrEmpty(textBox12.Text))
                Process.Start(textBox12.Text);
            if (!string.IsNullOrEmpty(textBox11.Text))
                Process.Start(textBox11.Text);
            if (!string.IsNullOrEmpty(textBox10.Text))
                Process.Start("explorer.exe", textBox10.Text);

            // 2. 关闭设置的应用
            if (!string.IsNullOrEmpty(textBox9.Text))
            {
                string appName = Path.GetFileNameWithoutExtension(textBox9.Text);
                CloseApplication(appName);
            }

            // 3. 恢复小说和视频
            if (checkBox4.Checked && !string.IsNullOrEmpty(settings.NovelPath))
            {
                string novelExePath = Path.Combine(Application.StartupPath, "Novel reading function.exe");
                if (!File.Exists(novelExePath))
                {
                    novelExePath = Path.Combine(Application.StartupPath, "Novel Reading Function", "Novel reading function.exe");
                }

                if (File.Exists(novelExePath))
                {
                    Process.Start(novelExePath, $"\"{settings.NovelPath}\"");
                }
                else
                {
                    MessageBox.Show($"未找到小说阅读程序！路径: {novelExePath}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            if (checkBox5.Checked && !string.IsNullOrEmpty(settings.VideoPath))
            {
                string videoExePath = Path.Combine(Application.StartupPath, "Video Watching function.exe");
                if (!File.Exists(videoExePath))
                {
                    videoExePath = Path.Combine(Application.StartupPath, "Video Watching Function", "Video Watching function.exe");
                }

                if (File.Exists(videoExePath))
                {
                    Process.Start(videoExePath, $"\"{settings.VideoPath}\"");
                }
                else
                {
                    MessageBox.Show($"未找到视频观看程序！路径: {videoExePath}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // 4. 恢复之前最小化的应用
            if (checkBox6.Checked)
            {
                RestoreMinimizedApplications();
            }

            // 5. 恢复静音状态
            if (checkBox1.Checked)
            {
                MuteSystem(wasMuted);
            }

            // 恢复主窗口
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.BringToFront();
        }

        private void CloseApplication(string appName)
        {
            try
            {
                var processes = Process.GetProcesses()
                    .Where(p => p.ProcessName.Equals(appName, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                foreach (Process process in processes)
                {
                    try
                    {
                        // 尝试优雅关闭
                        if (!process.CloseMainWindow())
                        {
                            // 如果优雅关闭失败，强制终止
                            process.Kill();
                        }
                        process.WaitForExit(100); // 等待最多100毫秒
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"关闭应用 {appName} 失败: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"查找应用 {appName} 失败: {ex.Message}");
            }
        }

        private void MinimizeApplication(string appName)
        {
            try
            {
                // 获取所有进程
                Process[] processes = Process.GetProcessesByName(appName);
                foreach (Process process in processes)
                {
                    // 获取主窗口句柄
                    if (process.MainWindowHandle != IntPtr.Zero)
                    {
                        // 最小化窗口
                        ShowWindow(process.MainWindowHandle, SW_MINIMIZE);
                        // 记录最小化的窗口
                        minimizedWindows.Add(process.MainWindowHandle);
                    }
                }

                // 如果没找到，尝试使用完整路径
                if (processes.Length == 0 && File.Exists(appName))
                {
                    string fileName = Path.GetFileNameWithoutExtension(appName);
                    MinimizeApplication(fileName);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"最小化应用错误: {ex.Message}");
            }
        }

        private void RestoreMinimizedApplications()
        {
            try
            {
                // 复制一份列表，避免在遍历时修改集合
                var windowsToRestore = new List<IntPtr>(minimizedWindows);
                minimizedWindows.Clear();

                foreach (IntPtr hWnd in windowsToRestore)
                {
                    if (hWnd != IntPtr.Zero && IsWindow(hWnd))
                    {
                        // 恢复窗口
                        ShowWindow(hWnd, SW_RESTORE);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"恢复应用错误: {ex.Message}");
            }
        }


        private bool IsSystemMuted()
        {
            try
            {
                uint volume;
                waveOutGetVolume(IntPtr.Zero, out volume);
                return volume == 0;
            }
            catch
            {
                return false;
            }
        }

        private void MuteSystem(bool mute)
        {
            try
            {
                // 如果已经是目标状态，不需要操作
                if (IsSystemMuted() == mute) return;

                // 模拟按下静音键
                keybd_event(VK_VOLUME_MUTE, 0, KEYEVENTF_EXTENDEDKEY, 0);
                keybd_event(VK_VOLUME_MUTE, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"静音操作错误: {ex.Message}");
            }
        }

        #region 事件处理
        private void Button1_Click(object sender, EventArgs e)
        {
            // 更改老板键
            using (var form = new BossKeyForm(settings.BossKey))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    // 尝试注册新热键
                    HotkeySetting oldKey = settings.BossKey;
                    settings.BossKey = form.SelectedHotkey;

                    try
                    {
                        RegisterBossKey();
                        label1.Text = "当前老板键：" + FormatHotkey(settings.BossKey);
                        SaveSettings();
                    }
                    catch
                    {
                        // 注册失败，恢复原来的热键
                        settings.BossKey = oldKey;
                        RegisterBossKey();
                        MessageBox.Show("新老板键注册失败，已恢复原来的设置。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private string FormatHotkey(HotkeySetting hotkey)
        {
            if (hotkey == null) return "未设置";

            List<string> parts = new List<string>();
            if (hotkey.Ctrl) parts.Add("Ctrl");
            if (hotkey.Alt) parts.Add("Alt");
            if (hotkey.Shift) parts.Add("Shift");
            parts.Add(hotkey.Key.ToString());

            return string.Join(" + ", parts);
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            // 浏览小说路径
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.SelectedPath = settings.NovelPath; // 设置初始路径
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    textBox1.Text = dialog.SelectedPath;
                    settings.NovelPath = dialog.SelectedPath;
                    SaveSettings();
                }
            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            // 浏览视频路径
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.SelectedPath = settings.VideoPath; // 设置初始路径
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    textBox2.Text = dialog.SelectedPath;
                    settings.VideoPath = dialog.SelectedPath;
                    SaveSettings();
                }
            }
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            // 浏览打开的应用 (老板来时)
            if (checkBox7.Checked)
            {
                BrowseInstalledSoftware(textBox3, settings.OnBossComing.OpenApps);
            }
            else
            {
                using (OpenFileDialog dialog = new OpenFileDialog())
                {
                    dialog.Filter = "应用程序|*.exe";
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        textBox3.Text = dialog.FileName;
                        if (settings.OnBossComing.OpenApps.Count == 0)
                            settings.OnBossComing.OpenApps.Add(dialog.FileName);
                        else
                            settings.OnBossComing.OpenApps[0] = dialog.FileName;
                        SaveSettings();
                    }
                }
            }
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            // 浏览打开的文件 (老板来时)
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    textBox4.Text = dialog.FileName;
                    if (settings.OnBossComing.OpenFiles.Count == 0)
                        settings.OnBossComing.OpenFiles.Add(dialog.FileName);
                    else
                        settings.OnBossComing.OpenFiles[0] = dialog.FileName;
                    SaveSettings();
                }
            }
        }

        private void Button8_Click(object sender, EventArgs e)
        {
            // 浏览关闭的应用 (老板来时)
            if (checkBox7.Checked)
            {
                BrowseInstalledSoftwareForClose(textBox8, settings.OnBossComing.CloseApps);
            }
            else
            {
                using (OpenFileDialog dialog = new OpenFileDialog())
                {
                    dialog.Filter = "应用程序|*.exe";
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        textBox8.Text = dialog.FileName;
                        if (settings.OnBossComing.CloseApps.Count == 0)
                            settings.OnBossComing.CloseApps.Add(dialog.FileName);
                        else
                            settings.OnBossComing.CloseApps[0] = dialog.FileName;
                        SaveSettings();
                    }
                }
            }
        }

        private void Button7_Click(object sender, EventArgs e)
        {
            // 浏览最小化的应用
            if (checkBox7.Checked)
            {
                BrowseInstalledSoftware(textBox6, settings.OnBossComing.MinimizeApps);
            }
            else
            {
                using (OpenFileDialog dialog = new OpenFileDialog())
                {
                    dialog.Filter = "应用程序|*.exe";
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        textBox6.Text = dialog.FileName;
                        if (settings.OnBossComing.MinimizeApps.Count == 0)
                            settings.OnBossComing.MinimizeApps.Add(dialog.FileName);
                        else
                            settings.OnBossComing.MinimizeApps[0] = dialog.FileName;
                        SaveSettings();
                    }
                }
            }
        }

        private void Button10_Click(object sender, EventArgs e)
        {
            // 浏览关闭的应用 (老板走后)
            if (checkBox7.Checked)
            {
                BrowseInstalledSoftware(textBox9, settings.OnBossLeaving.CloseApps);
            }
            else
            {
                using (OpenFileDialog dialog = new OpenFileDialog())
                {
                    dialog.Filter = "应用程序|*.exe";
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        textBox9.Text = dialog.FileName;
                        if (settings.OnBossLeaving.CloseApps.Count == 0)
                            settings.OnBossLeaving.CloseApps.Add(dialog.FileName);
                        else
                            settings.OnBossLeaving.CloseApps[0] = dialog.FileName;
                        SaveSettings();
                    }
                }
            }
        }

        private void Button12_Click(object sender, EventArgs e)
        {
            // 浏览打开的应用 (老板走后)
            if (checkBox7.Checked)
            {
                BrowseInstalledSoftware(textBox12, settings.OnBossLeaving.OpenApps);
            }
            else
            {
                using (OpenFileDialog dialog = new OpenFileDialog())
                {
                    dialog.Filter = "应用程序|*.exe";
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        textBox12.Text = dialog.FileName;
                        if (settings.OnBossLeaving.OpenApps.Count == 0)
                            settings.OnBossLeaving.OpenApps.Add(dialog.FileName);
                        else
                            settings.OnBossLeaving.OpenApps[0] = dialog.FileName;
                        SaveSettings();
                    }
                }
            }
        }

        private void Button11_Click(object sender, EventArgs e)
        {
            // 浏览打开的文件 (老板走后)
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    textBox11.Text = dialog.FileName;
                    if (settings.OnBossLeaving.OpenFiles.Count == 0)
                        settings.OnBossLeaving.OpenFiles.Add(dialog.FileName);
                    else
                        settings.OnBossLeaving.OpenFiles[0] = dialog.FileName;
                    SaveSettings();
                }
            }
        }

        private void BrowseInstalledSoftware(TextBox targetTextBox, List<string> targetList)
        {
            using (var form = new InstalledSoftwareForm())
            {
                if (form.ShowDialog() == DialogResult.OK && !string.IsNullOrEmpty(form.SelectedPath))
                {
                    targetTextBox.Text = form.SelectedPath;

                    // 直接替换而不是判断Count
                    if (targetList.Count > 0) targetList[0] = form.SelectedPath;
                    else targetList.Add(form.SelectedPath);

                    SaveSettings();
                }
            }
        }

        private void BrowseInstalledSoftwareForClose(TextBox targetTextBox, List<string> targetList)
        {
            using (var form = new InstalledSoftwareForm())
            {
                if (form.ShowDialog() == DialogResult.OK && !string.IsNullOrEmpty(form.SelectedPath))
                {
                    // 对于关闭操作，我们只需要进程名
                    string processName = Path.GetFileNameWithoutExtension(form.SelectedPath);
                    targetTextBox.Text = processName;
                    if (targetList.Count == 0)
                        targetList.Add(processName);
                    else
                        targetList[0] = processName;
                    SaveSettings();
                }
            }
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            // 检查小说和视频程序是否存在
            string novelExePath = Path.Combine(Application.StartupPath, "Novel reading function.exe");
            if (!File.Exists(novelExePath))
            {
                novelExePath = Path.Combine(Application.StartupPath, "Novel Reading Function", "Novel reading function.exe");
            }

            string videoExePath = Path.Combine(Application.StartupPath, "Video Watching function.exe");
            if (!File.Exists(videoExePath))
            {
                videoExePath = Path.Combine(Application.StartupPath, "Video Watching Function", "Video Watching function.exe");
            }

            if (!File.Exists(novelExePath) || !File.Exists(videoExePath))
            {
                MessageBox.Show("未找到小说或视频功能程序，请确保程序完整安装！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 重置统计信息（如果是新的一天）
            if (settings.LastFishingDate.Date != DateTime.Today)
            {
                settings.BossVisitCount = 0;
                settings.FishingDuration = TimeSpan.Zero;
                settings.LastFishingDate = DateTime.Today;
            }

            // 开始摸鱼
            SaveSettings();
            this.Hide();
            trayIcon.Visible = true;
            trayIcon.ShowBalloonTip(1000, "摸鱼助手", "已开始运行，按老板键切换状态", ToolTipIcon.Info);
            isInFishingMode = true; // 进入摸鱼模式

            // 记录摸鱼开始时间
            fishingStartTime = DateTime.Now;
            fishingTimer.Start();

            // 如果设置了摸鱼时打开小说
            if (settings.OpenNovelWhenFishing && !string.IsNullOrEmpty(settings.NovelPath))
            {
                OpenNovel();
            }

            // 如果设置了摸鱼时打开视频
            if (settings.OpenVideoWhenFishing && !string.IsNullOrEmpty(settings.VideoPath))
            {
                OpenVideo();
            }
        }

        private void OpenNovel()
        {
            try
            {
                string novelExePath = Path.Combine(Application.StartupPath, "Novel reading function.exe");
                if (!File.Exists(novelExePath))
                {
                    novelExePath = Path.Combine(Application.StartupPath, "Novel Reading Function", "Novel reading function.exe");
                }

                if (File.Exists(novelExePath))
                {
                    Process.Start(novelExePath, $"\"{settings.NovelPath}\"");
                }
                else
                {
                    MessageBox.Show($"未找到小说阅读程序！路径: {novelExePath}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开小说失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OpenVideo()
        {
            try
            {
                string videoExePath = Path.Combine(Application.StartupPath, "Video Watching function.exe");
                if (!File.Exists(videoExePath))
                {
                    videoExePath = Path.Combine(Application.StartupPath, "Video Watching Function", "Video Watching function.exe");
                }

                if (File.Exists(videoExePath))
                {
                    Process.Start(videoExePath, $"\"{settings.VideoPath}\"");
                }
                else
                {
                    MessageBox.Show($"未找到视频观看程序！路径: {videoExePath}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开视频失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        protected override void OnResize(EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                this.Hide();
                trayIcon.Visible = true;
                trayIcon.ShowBalloonTip(1000, "摸鱼助手", "程序已最小化到系统托盘", ToolTipIcon.Info);
            }

            base.OnResize(e);
        }

        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 打开GitHub仓库
            Process.Start("https://github.com/SYSTEM-MEMZ-XEK/Fishing-assistant.git");
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            fishingTimer.Stop(); // 停止计时器
            UnregisterBossKey();
            SaveSettings();

            if (trayIcon != null)
            {
                trayIcon.Visible = false;
                trayIcon.Dispose();
            }
        }

        private void TextBox7_Validating(object sender, CancelEventArgs e)
        {
            if (ParseSize(textBox7.Text, out Size size))
            {
                settings.NovelWindowSize = size;
                SaveSettings();
            }
            else
            {
                MessageBox.Show("请输入有效的窗口尺寸（例如：400*100）");
                textBox7.Text = $"{settings.NovelWindowSize.Width}*{settings.NovelWindowSize.Height}";
            }
        }

        private void TextBox13_Validating(object sender, CancelEventArgs e)
        {
            if (ParseSize(textBox13.Text, out Size size))
            {
                settings.VideoWindowSize = size;
                SaveSettings();
            }
            else
            {
                MessageBox.Show("请输入有效的窗口尺寸（例如：400*400）");
                textBox13.Text = $"{settings.VideoWindowSize.Width}*{settings.VideoWindowSize.Height}";
            }
        }

        private bool ParseSize(string text, out Size size)
        {
            size = Size.Empty;
            // 支持多种分隔符
            char[] separators = { '*', 'x', 'X', '×' };
            string[] parts = text.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length != 2) return false;

            if (int.TryParse(parts[0], out int width) &&
                int.TryParse(parts[1], out int height))
            {
                size = new Size(width, height);
                return true;
            }
            return false;
        }

        // 更改字体按钮
        private void Button9_Click(object sender, EventArgs e)
        {
            try
            {
                using (FontDialog fontDialog = new FontDialog())
                {
                    fontDialog.Font = settings.NovelFont;
                    if (fontDialog.ShowDialog() == DialogResult.OK)
                    {
                        settings.NovelFont = fontDialog.Font;
                        label18.Text = $"当前文本字体: {fontDialog.Font.Name}, {fontDialog.Font.Size}pt";
                        SaveSettings();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置字体时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 更改颜色按钮
        private void Button13_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                colorDialog.Color = settings.NovelTextColor;
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    settings.NovelTextColor = colorDialog.Color;
                    label21.Text = $"当前文本颜色: {colorDialog.Color.Name}";
                    SaveSettings();
                }
            }
        }

        // 小说窗口透明
        private void CheckBox8_CheckedChanged(object sender, EventArgs e)
        {
            settings.NovelWindowTransparent = checkBox8.Checked;
            SaveSettings();
        }

        // 自动翻页
        private void CheckBox9_CheckedChanged(object sender, EventArgs e)
        {
            settings.NovelAutoTurnPage = checkBox9.Checked;
            SaveSettings();
        }

        // 翻页时长
        private void NumericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            settings.NovelTurnPageSeconds = (int)numericUpDown1.Value;
            SaveSettings();
        }

        // 视频自动切换
        private void CheckBox10_CheckedChanged(object sender, EventArgs e)
        {
            settings.VideoAutoSwitch = checkBox10.Checked;
            SaveSettings();
        }
        #endregion

        #region 设置类
        public class FishingSettings
        {
            public HotkeySetting BossKey { get; set; } = new HotkeySetting
            {
                Ctrl = true,
                Alt = true,
                Key = Keys.F12
            };

            public string NovelPath { get; set; } = "";
            public string VideoPath { get; set; } = "";

            public BossAction OnBossComing { get; set; } = new BossAction();
            public BossAction OnBossLeaving { get; set; } = new BossAction();

            // 新增小说设置
            [JsonIgnore]
            public Font NovelFont { get; set; } = new Font("宋体", 12);

            public string NovelFontSerialized
            {
                get => FontToString(NovelFont);
                set => NovelFont = StringToFont(value);
            }

            [JsonIgnore]
            public Color NovelTextColor { get; set; } = Color.Black;

            public string NovelTextColorHex
            {
                get => ColorTranslator.ToHtml(NovelTextColor);
                set => NovelTextColor = ColorTranslator.FromHtml(value);
            }

            public bool NovelWindowTransparent { get; set; } = false;
            public bool NovelAutoTurnPage { get; set; } = false;
            public int NovelTurnPageSeconds { get; set; } = 10;
            public Size NovelWindowSize { get; set; } = new Size(400, 100);

            // 新增视频设置
            public bool VideoAutoSwitch { get; set; } = false;
            public Size VideoWindowSize { get; set; } = new Size(400, 400);

            // 辅助方法：将字体转换为字符串
            private string FontToString(Font font)
            {
                return $"{font.Name}|{font.Size}|{(int)font.Style}";
            }

            // 辅助方法：将字符串转换为字体
            private Font StringToFont(string fontString)
            {
                string[] parts = fontString.Split('|');
                if (parts.Length != 3) return new Font("宋体", 12);

                return new Font(parts[0],
                              float.Parse(parts[1]),
                              (FontStyle)int.Parse(parts[2]));
            }

            // 新增摸鱼时自动打开选项
            // 新增：摸鱼时打开小说
            public bool OpenNovelWhenFishing { get; set; } = false;

            // 新增：摸鱼时打开视频
            public bool OpenVideoWhenFishing { get; set; } = false;

            // 新增：视频全屏模式
            public bool VideoFullScreen { get; set; } = false;

            // 新增：统计信息
            public int BossVisitCount { get; set; } = 0;
            public TimeSpan FishingDuration { get; set; } = TimeSpan.Zero;
            public DateTime LastFishingDate { get; set; } = DateTime.MinValue;
        }

        public class HotkeySetting
        {
            public bool Ctrl { get; set; } = false;
            public bool Alt { get; set; } = false;
            public bool Shift { get; set; } = false;
            public Keys Key { get; set; } = Keys.None;
        }

        public class BossAction
        {
            public List<string> OpenApps { get; set; } = new List<string>();
            public List<string> OpenFiles { get; set; } = new List<string>();
            public List<string> OpenWebs { get; set; } = new List<string>();

            public List<string> CloseApps { get; set; } = new List<string>();
            public bool CloseNovel { get; set; } = false;
            public bool CloseVideo { get; set; } = false;

            public List<string> MinimizeApps { get; set; } = new List<string>();
            public bool MuteWhenMinimize { get; set; } = false;

            public bool RestoreNovel { get; set; } = false;
            public bool RestoreVideo { get; set; } = false;
            public bool RestoreClosedApps { get; set; } = false;
        }
        #endregion
    }

    #region 老板键设置窗体
    public class BossKeyForm : Form
    {
        public HotkeySetting SelectedHotkey { get; private set; } = new HotkeySetting();
        private Label lblCurrentKey;
        private Button btnCancel;
        private Button btnClear;
        private Button btnOK;

        public BossKeyForm(HotkeySetting currentHotkey)
        {
            this.Size = new Size(400, 200);
            this.Text = "设置老板键";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.KeyPreview = true;

            // 初始化控件
            lblCurrentKey = new Label
            {
                Text = "当前组合键: " + FormatHotkey(currentHotkey),
                Location = new Point(20, 20),
                AutoSize = true
            };

            Label lblInfo = new Label
            {
                Text = "请按下新的组合键 (Ctrl, Alt, Shift + 其他键):",
                Location = new Point(20, 50),
                AutoSize = true
            };

            btnClear = new Button
            {
                Text = "清除",
                Location = new Point(20, 100),
                Size = new Size(80, 30)
            };

            btnCancel = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                Location = new Point(110, 100),
                Size = new Size(80, 30)
            };

            btnOK = new Button
            {
                Text = "确定",
                DialogResult = DialogResult.OK,
                Location = new Point(200, 100),
                Size = new Size(80, 30)
            };

            // 事件处理
            btnClear.Click += (s, e) => {
                // 确保currentHotkey不为null
                if (currentHotkey == null)
                {
                    currentHotkey = new HotkeySetting
                    {
                        Ctrl = true,
                        Alt = true,
                        Key = Keys.F12
                    };
                }
                SelectedHotkey = new HotkeySetting();
                lblCurrentKey.Text = "当前组合键: 未设置";
            };

            btnOK.Click += (s, e) => {
                if (SelectedHotkey.Key == Keys.None)
                {
                    MessageBox.Show("请设置有效的组合键", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.DialogResult = DialogResult.None;
                }
            };

            this.KeyDown += BossKeyForm_KeyDown;

            // 添加控件
            this.Controls.Add(lblCurrentKey);
            this.Controls.Add(lblInfo);
            this.Controls.Add(btnClear);
            this.Controls.Add(btnCancel);
            this.Controls.Add(btnOK);

            // 设置当前热键
            SelectedHotkey = currentHotkey ?? new HotkeySetting();
        }

        private void BossKeyForm_KeyDown(object sender, KeyEventArgs e)
        {
            // 忽略修饰键本身
            if (e.KeyCode == Keys.ControlKey || e.KeyCode == Keys.ShiftKey || e.KeyCode == Keys.Menu)
                return;

            // 设置热键
            SelectedHotkey.Ctrl = e.Control;
            SelectedHotkey.Alt = e.Alt;
            SelectedHotkey.Shift = e.Shift;
            SelectedHotkey.Key = e.KeyCode;

            // 更新显示
            lblCurrentKey.Text = "当前组合键: " + FormatHotkey(SelectedHotkey);
            e.Handled = true;
        }

        private string FormatHotkey(HotkeySetting hotkey)
        {
            if (hotkey == null || hotkey.Key == Keys.None) return "未设置";

            List<string> parts = new List<string>();
            if (hotkey.Ctrl) parts.Add("Ctrl");
            if (hotkey.Alt) parts.Add("Alt");
            if (hotkey.Shift) parts.Add("Shift");
            parts.Add(hotkey.Key.ToString());

            return string.Join(" + ", parts);
        }
    }
    #endregion

    #region 已安装软件浏览窗体
    public class InstalledSoftwareForm : Form
    {
        public string SelectedPath { get; private set; }

        public InstalledSoftwareForm()
        {
            this.Font = new Font("微软雅黑", 9F, FontStyle.Regular, GraphicsUnit.Point);
            InitializeComponent();
            LoadInstalledSoftware();
            this.listViewSoftware.Font = this.Font;
            this.btnSelect.Font = this.Font;
        }

        private void InitializeComponent()
        {
            this.Text = "已安装软件";
            this.Size = new Size(600, 400);

            // 创建列表视图
            listViewSoftware = new ListView
            {
                View = View.Details,
                FullRowSelect = true,
                Dock = DockStyle.Fill,
                MultiSelect = false
            };
            listViewSoftware.Columns.Add("软件名称", 200);
            listViewSoftware.Columns.Add("版本", 100);
            listViewSoftware.Columns.Add("安装路径", 300);
            listViewSoftware.DoubleClick += ListViewSoftware_DoubleClick;

            // 创建选择按钮
            btnSelect = new Button
            {
                Text = "选择",
                Dock = DockStyle.Bottom,
                Height = 30
            };
            btnSelect.Click += BtnSelect_Click;

            // 添加控件
            this.Controls.Add(listViewSoftware);
            this.Controls.Add(btnSelect);
        }

        private void ListViewSoftware_DoubleClick(object sender, EventArgs e)
        {
            BtnSelect_Click(sender, e);
        }

        private void LoadInstalledSoftware()
        {
            // 获取已安装软件列表
            var softwareList = GetInstalledSoftware();

            foreach (var software in softwareList)
            {
                var item = new ListViewItem(software.DisplayName);
                item.SubItems.Add(software.DisplayVersion);
                item.SubItems.Add(software.InstallLocation);
                item.Tag = software.ExecutablePath ?? software.InstallLocation; // 存储可执行路径或安装目录
                listViewSoftware.Items.Add(item);
            }
        }

        private List<SoftwareInfo> GetInstalledSoftware()
        {
            var softwareList = new List<SoftwareInfo>();
            var uniqueNames = new HashSet<string>();

            string[] registryKeys = {
        @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        @"SOFTWARE\Classes\Installer\Products"
    };

            try
            {
                // 检查本地机器注册表
                foreach (var keyPath in registryKeys)
                {
                    using (RegistryKey key = Registry.LocalMachine.OpenSubKey(keyPath))
                    {
                        if (key == null) continue;

                        foreach (string subkeyName in key.GetSubKeyNames())
                        {
                            try
                            {
                                using (RegistryKey subkey = key.OpenSubKey(subkeyName))
                                {
                                    string displayName = subkey.GetValue("DisplayName") as string;
                                    if (string.IsNullOrEmpty(displayName)) continue;

                                    string displayVersion = subkey.GetValue("DisplayVersion") as string ?? "";
                                    string publisher = subkey.GetValue("Publisher") as string ?? "";
                                    string installLocation = subkey.GetValue("InstallLocation") as string;
                                    string executablePath = subkey.GetValue("DisplayIcon") as string;
                                    string uninstallString = subkey.GetValue("UninstallString") as string;

                                    // 清理可执行路径
                                    if (!string.IsNullOrEmpty(executablePath))
                                    {
                                        // 处理带引号的路径
                                        if (executablePath.Contains("\""))
                                        {
                                            executablePath = executablePath.Split('\"')[1];
                                        }
                                        // 处理带逗号的路径（通常是图标索引）
                                        else if (executablePath.Contains(","))
                                        {
                                            executablePath = executablePath.Split(',')[0];
                                        }
                                        // 处理带空格的长路径
                                        else if (executablePath.Contains(" "))
                                        {
                                            executablePath = executablePath.Split(' ')[0];
                                        }
                                    }

                                    // 处理安装路径为空的情况
                                    if (string.IsNullOrEmpty(installLocation))
                                    {
                                        if (!string.IsNullOrEmpty(uninstallString))
                                        {
                                            // 尝试从卸载字符串中提取路径
                                            if (uninstallString.Contains("\""))
                                            {
                                                installLocation = Path.GetDirectoryName(uninstallString.Split('\"')[1]);
                                            }
                                            else if (uninstallString.Contains(".exe"))
                                            {
                                                installLocation = Path.GetDirectoryName(uninstallString);
                                            }
                                        }
                                        else if (!string.IsNullOrEmpty(executablePath))
                                        {
                                            installLocation = Path.GetDirectoryName(executablePath);
                                        }
                                    }

                                    // 确保有有效路径
                                    if (string.IsNullOrEmpty(installLocation) &&
                                        string.IsNullOrEmpty(executablePath)) continue;

                                    // 使用名称+版本+发布者作为唯一标识
                                    string uniqueId = $"{displayName}|{displayVersion}|{publisher}";
                                    if (uniqueNames.Contains(uniqueId)) continue;

                                    softwareList.Add(new SoftwareInfo
                                    {
                                        DisplayName = displayName,
                                        DisplayVersion = displayVersion,
                                        Publisher = publisher,
                                        InstallLocation = installLocation ?? "",
                                        ExecutablePath = executablePath ?? ""
                                    });

                                    uniqueNames.Add(uniqueId);
                                }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"读取注册表子键 {subkeyName} 失败: {ex.Message}");
                            }
                        }
                    }
                }

                // 检查当前用户注册表
                using (RegistryKey userKey = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"))
                {
                    if (userKey != null)
                    {
                        foreach (string subkeyName in userKey.GetSubKeyNames())
                        {
                            try
                            {
                                using (RegistryKey subkey = userKey.OpenSubKey(subkeyName))
                                {
                                    string displayName = subkey.GetValue("DisplayName") as string;
                                    if (string.IsNullOrEmpty(displayName)) continue;

                                    string displayVersion = subkey.GetValue("DisplayVersion") as string ?? "";
                                    string publisher = subkey.GetValue("Publisher") as string ?? "";
                                    string installLocation = subkey.GetValue("InstallLocation") as string;
                                    string executablePath = subkey.GetValue("DisplayIcon") as string;
                                    string uninstallString = subkey.GetValue("UninstallString") as string;

                                    // 清理可执行路径（同上）
                                    if (!string.IsNullOrEmpty(executablePath))
                                    {
                                        if (executablePath.Contains("\""))
                                        {
                                            executablePath = executablePath.Split('\"')[1];
                                        }
                                        else if (executablePath.Contains(","))
                                        {
                                            executablePath = executablePath.Split(',')[0];
                                        }
                                        else if (executablePath.Contains(" "))
                                        {
                                            executablePath = executablePath.Split(' ')[0];
                                        }
                                    }

                                    // 处理安装路径
                                    if (string.IsNullOrEmpty(installLocation))
                                    {
                                        if (!string.IsNullOrEmpty(uninstallString))
                                        {
                                            if (uninstallString.Contains("\""))
                                            {
                                                installLocation = Path.GetDirectoryName(uninstallString.Split('\"')[1]);
                                            }
                                            else if (uninstallString.Contains(".exe"))
                                            {
                                                installLocation = Path.GetDirectoryName(uninstallString);
                                            }
                                        }
                                        else if (!string.IsNullOrEmpty(executablePath))
                                        {
                                            installLocation = Path.GetDirectoryName(executablePath);
                                        }
                                    }

                                    if (string.IsNullOrEmpty(installLocation) &&
                                        string.IsNullOrEmpty(executablePath)) continue;

                                    string uniqueId = $"{displayName}|{displayVersion}|{publisher}";
                                    if (uniqueNames.Contains(uniqueId)) continue;

                                    softwareList.Add(new SoftwareInfo
                                    {
                                        DisplayName = displayName,
                                        DisplayVersion = displayVersion,
                                        Publisher = publisher,
                                        InstallLocation = installLocation ?? "",
                                        ExecutablePath = executablePath ?? ""
                                    });

                                    uniqueNames.Add(uniqueId);
                                }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"读取用户注册表子键 {subkeyName} 失败: {ex.Message}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"读取注册表失败: {ex.Message}");
            }

            // 后处理：确保路径格式正确
            foreach (var software in softwareList)
            {
                try
                {
                    // 标准化路径
                    if (!string.IsNullOrEmpty(software.InstallLocation))
                    {
                        // 检查路径是否存在
                        if (Directory.Exists(software.InstallLocation))
                        {
                            software.InstallLocation = Path.GetFullPath(software.InstallLocation);
                        }
                        else
                        {
                            Debug.WriteLine($"安装目录不存在: {software.InstallLocation}");
                            software.InstallLocation = ""; // 清除非法路径
                        }
                    }

                    if (!string.IsNullOrEmpty(software.ExecutablePath))
                    {
                        // 检查文件是否存在
                        if (File.Exists(software.ExecutablePath))
                        {
                            software.ExecutablePath = Path.GetFullPath(software.ExecutablePath);
                        }
                        else
                        {
                            Debug.WriteLine($"可执行文件不存在: {software.ExecutablePath}");
                            software.ExecutablePath = ""; // 清除非法路径
                        }
                    }

                    // 如果可执行路径为空但安装位置有值
                    if (string.IsNullOrEmpty(software.ExecutablePath) &&
                        !string.IsNullOrEmpty(software.InstallLocation) &&
                        Directory.Exists(software.InstallLocation))
                    {
                        try
                        {
                            // 尝试在安装目录中查找可执行文件
                            var exeFiles = Directory.GetFiles(software.InstallLocation, "*.exe", SearchOption.TopDirectoryOnly);
                            if (exeFiles.Length > 0)
                            {
                                // 只取第一个找到的exe文件
                                software.ExecutablePath = exeFiles[0];
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"在目录 {software.InstallLocation} 中查找可执行文件失败: {ex.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"处理软件 {software.DisplayName} 时出错: {ex.Message}");
                }
            }

            return softwareList
                .OrderBy(s => s.DisplayName)
                .ToList();
        }

        private void BtnSelect_Click(object sender, EventArgs e)
        {
            if (listViewSoftware.SelectedItems.Count > 0)
            {
                SelectedPath = listViewSoftware.SelectedItems[0].Tag as string;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                MessageBox.Show("请选择一个软件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private ListView listViewSoftware;
        private Button btnSelect;
    }

    public class SoftwareInfo
    {
        public string DisplayName { get; set; } = "";
        public string DisplayVersion { get; set; } = "";
        public string Publisher { get; set; } = "";
        public string InstallLocation { get; set; } = "";
        public string ExecutablePath { get; set; } = "";
    }
    #endregion
}
    public class SoftwareInfo
    {
        public string DisplayName { get; set; }
        public string DisplayVersion { get; set; }
        public string InstallLocation { get; set; }
        public string ExecutablePath { get; set; }
    }


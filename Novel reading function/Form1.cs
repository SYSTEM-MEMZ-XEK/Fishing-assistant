using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace NovelReader
{
    public class NovelReaderForm : Form
    {
        private List<string> novelFiles = new List<string>();
        private int currentNovelIndex = 0;
        private int currentPage = 0;
        private System.Windows.Forms.Timer autoTurnTimer;
        private FishingSettings settings = new FishingSettings();
        private TextBox txtContent = new TextBox();

        // 窗口拖动API
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();

        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HT_CAPTION = 0x2;

        public NovelReaderForm(string novelPath)
        {
            InitializeControls();
            LoadSettings();
            LoadNovels(novelPath);
            InitializeTimer();
            ApplySettings();
        }

        private void InitializeControls()
        {
            // 使用洋红色作为透明键
            Color transparentColor = Color.Magenta;

            // 初始化文本控件
            txtContent = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.None, // 无滚动条
                Font = new Font("微软雅黑", 10),
                BackColor = transparentColor, // 背景设为洋红色（透明）
                ForeColor = Color.White,      // 文本为白色
                BorderStyle = BorderStyle.None,
                Cursor = Cursors.Arrow        // 箭头光标
            };

            // 添加鼠标拖动支持
            txtContent.MouseDown += (s, e) =>
            {
                if (e.Button == MouseButtons.Left)
                {
                    ReleaseCapture();
                    SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
                }
            };

            // 添加Enter事件处理程序来移除焦点
            txtContent.Enter += (s, e) => this.Focus();

            this.Controls.Add(txtContent);

            // 窗体设置
            this.Text = "小说阅读器";
            this.Size = new Size(500, 300);
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.KeyPreview = true;
            this.KeyDown += MainForm_KeyDown;
            this.BackColor = transparentColor;        // 窗体背景设为洋红色
            this.TransparencyKey = transparentColor;  // 透明键设为洋红色
            this.TopMost = true;

            // 添加此事件确保窗体加载时文本框没有焦点
            this.Shown += (s, e) => this.Focus();
        }

        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            // 空格键翻页
            if (e.KeyCode == Keys.Space)
            {
                NextPage();
                e.Handled = true;
            }
            // ESC键退出
            else if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
            // N键下一本小说
            else if (e.KeyCode == Keys.N)
            {
                currentNovelIndex = (currentNovelIndex + 1) % novelFiles.Count;
                LoadNovelContent(currentNovelIndex);
                e.Handled = true;
            }
            // P键上一本小说
            else if (e.KeyCode == Keys.P)
            {
                currentNovelIndex = (currentNovelIndex - 1 + novelFiles.Count) % novelFiles.Count;
                LoadNovelContent(currentNovelIndex);
                e.Handled = true;
            }
            // F键切换窗口置顶
            else if (e.KeyCode == Keys.F)
            {
                this.TopMost = !this.TopMost;
            }
        }

        private void LoadSettings()
        {
            // 修改：配置文件位于应用程序的上一级目录
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

        private void LoadNovels(string novelPath)
        {
            if (string.IsNullOrEmpty(novelPath))
            {
                novelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Novels");
            }

            if (!Directory.Exists(novelPath))
            {
                Directory.CreateDirectory(novelPath);
            }

            novelFiles = Directory.GetFiles(novelPath, "*.txt")
                .Concat(Directory.GetFiles(novelPath, "*.doc"))
                .Concat(Directory.GetFiles(novelPath, "*.docx"))
                .ToList();

            if (novelFiles.Count > 0)
            {
                LoadNovelContent(0);
            }
            else
            {
                txtContent.Text = $"未找到小说文件: {novelPath}\n请将小说文件(.txt/.doc/.docx)放入此目录";
            }
        }

        private void LoadNovelContent(int index)
        {
            if (index < 0 || index >= novelFiles.Count) return;

            try
            {
                string content = File.ReadAllText(novelFiles[index]);
                ShowContent(content);
                currentNovelIndex = index;
                currentPage = 0;
                this.Text = $"小说阅读器 - {Path.GetFileName(novelFiles[index])}";
            }
            catch (Exception ex)
            {
                txtContent.Text = $"加载小说失败: {ex.Message}";
            }
        }

        private void ShowContent(string content)
        {
            txtContent.Text = content;
            txtContent.Select(0, 0);
            txtContent.ScrollToCaret();

            // 确保文本框没有焦点
            this.Focus();
        }

        private void InitializeTimer()
        {
            autoTurnTimer = new System.Windows.Forms.Timer();
            autoTurnTimer.Tick += AutoTurnPage;
            autoTurnTimer.Interval = settings.NovelTurnPageSeconds * 1000;
            if (settings.NovelAutoTurnPage) autoTurnTimer.Start();
        }

        private void AutoTurnPage(object sender, EventArgs e)
        {
            NextPage();
        }

        private void NextPage()
        {
            if (txtContent.TextLength == 0) return;

            int visibleLines = txtContent.ClientSize.Height / txtContent.Font.Height;
            int charsPerPage = visibleLines * (txtContent.Width / (int)(txtContent.Font.Size * 0.7));

            currentPage++;
            int startIndex = currentPage * charsPerPage;

            if (startIndex >= txtContent.TextLength)
            {
                // 小说结束，切换到下一本
                currentNovelIndex = (currentNovelIndex + 1) % novelFiles.Count;
                LoadNovelContent(currentNovelIndex);
                return;
            }

            txtContent.Select(startIndex, 0);
            txtContent.ScrollToCaret();
        }

        private void ApplySettings()
        {
            // 应用字体设置
            try
            {
                string[] fontParts = settings.NovelFontSerialized.Split('|');
                if (fontParts.Length >= 3)
                {
                    string fontName = fontParts[0];
                    float fontSize = float.Parse(fontParts[1]);
                    FontStyle fontStyle = (FontStyle)int.Parse(fontParts[2]);

                    txtContent.Font = new Font(fontName, fontSize, fontStyle);
                }
            }
            catch
            {
                // 使用默认字体
                txtContent.Font = new Font("微软雅黑", 10);
            }

            // 强制文本为纯白色（忽略设置中的颜色值）
            txtContent.ForeColor = Color.White;

            // 应用窗口大小
            try
            {
                if (!string.IsNullOrEmpty(settings.NovelWindowSize))
                {
                    // 使用正确的分隔符（可能是'*'或'x'）
                    char[] separators = { '*', 'x', 'X', '×' };
                    string[] sizeParts = settings.NovelWindowSize.Split(separators);

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
                this.Size = new Size(500, 300);
            }

            // 应用自动翻页设置
            if (autoTurnTimer != null)
            {
                autoTurnTimer.Interval = settings.NovelTurnPageSeconds * 1000;
                autoTurnTimer.Enabled = settings.NovelAutoTurnPage;
            }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            // 窗口大小变化时重置页码
            currentPage = 0;
            txtContent.Select(0, 0);
            txtContent.ScrollToCaret();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            autoTurnTimer?.Stop();
            base.OnFormClosing(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                autoTurnTimer?.Dispose();
                txtContent?.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    // 设置类
    public class FishingSettings
    {
        // 字体设置
        public string NovelFontSerialized { get; set; } = "微软雅黑|9|0";

        // 文本颜色
        public string NovelTextColorHex { get; set; } = "White";

        // 窗口透明
        public bool NovelWindowTransparent { get; set; } = false;

        // 自动翻页
        public bool NovelAutoTurnPage { get; set; } = false;

        // 翻页间隔（秒）
        public int NovelTurnPageSeconds { get; set; } = 10;

        // 窗口大小
        public string NovelWindowSize { get; set; } = "400,300";

        // 新增：摸鱼时打开小说选项
        public bool OpenNovelWhenFishing { get; set; } = false;

        // 新增：小说文件夹路径
        public string NovelPath { get; set; } = "";
    }
}
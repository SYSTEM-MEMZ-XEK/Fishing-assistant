using System;
using System.Windows.Forms;
using VideoPlayer;

namespace Video_Watching_function
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string videoPath = string.Empty;
            if (args.Length > 0)
            {
                videoPath = args[0];
            }
            else
            {
                // 如果没有提供参数，使用默认路径或让用户选择
                using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.Description = "请选择视频目录";
                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        videoPath = folderDialog.SelectedPath;
                    }
                    else
                    {
                        return; // 用户取消
                    }
                }
            }

            Application.Run(new VideoPlayerForm(videoPath));
        }
    }
}
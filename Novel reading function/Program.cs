using NovelReader;
using System;
using System.Windows.Forms;

namespace Novel_reading_function
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string novelPath = string.Empty;
            if (args.Length > 0)
            {
                novelPath = args[0];
            }
            else
            {
                // 如果没有提供参数，使用默认路径或让用户选择
                using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.Description = "请选择小说目录";
                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        novelPath = folderDialog.SelectedPath;
                    }
                    else
                    {
                        return; // 用户取消
                    }
                }
            }

            Application.Run(new NovelReaderForm(novelPath));
        }
    }
}
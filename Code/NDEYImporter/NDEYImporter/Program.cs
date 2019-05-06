using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace NDEYImporter
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            //初始化本地数据库
            NDEYImporter.DB.ConnectionManager.Open(System.IO.Path.Combine(Application.StartupPath, "static.db"));

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
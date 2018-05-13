using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;

namespace LTZN
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new 炼铁智能主窗体());
        }

        public static void Log(string content)
        {
            File.WriteAllText(@"c:\temp.txt", content);
        }
    }
}
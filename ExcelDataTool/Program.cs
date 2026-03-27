using System;
using System.Windows.Forms;

namespace ExcelDataTool
{
    internal static class Program
    {
        /// <summary>
        ///  应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            // 为了自定义更高清的 UI (适用于 .NET 6/8)
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }
    }
}
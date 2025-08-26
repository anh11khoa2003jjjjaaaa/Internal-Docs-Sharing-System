using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ChiaseNoiBo
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Thiết lập văn hóa Việt Nam
            //CultureInfo viCulture = new CultureInfo("vi-VN");
            //Thread.CurrentThread.CurrentCulture = viCulture;
            //Thread.CurrentThread.CurrentUICulture = viCulture;
            Application.Run(new Login());
        }
    }
}

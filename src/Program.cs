using System;
using System.Windows.Forms;

namespace VirtualDesktopIndicator
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            using (TrayIndicator ti = new TrayIndicator())
            {
                Application.Run();
            }
        }
    }
}

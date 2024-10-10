using System;
using System.Windows.Forms;

namespace DocumentAgentDesktop
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Global exception handling
            Application.ThreadException += new System.Threading.ThreadExceptionEventHandler(Application_ThreadException);
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);

            Application.Run(new frmMain());
        }

        private static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            // Handle UI thread exceptions
            MessageBox.Show(e.Exception.Message, "Unhandled UI Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            // Handle non-UI thread exceptions
            Exception ex = (Exception)e.ExceptionObject;
            MessageBox.Show(ex.Message, "Unhandled Non-UI Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}

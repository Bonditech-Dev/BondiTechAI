using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;


// Must run as administrator for this to work and access the config file. 
namespace ConfigureDocumentAgent
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

            DocumentAgent.ConfigureSettings TestForm = new DocumentAgent.ConfigureSettings();
            TestForm.ConfigurationFile = Properties.Settings.Default.ConfigFile; //AppDomain.CurrentDomain.SetupInformation.ConfigurationFile;
            TestForm.LoadSettings();
            TestForm.ShowDialog();

            //Application.Run(new Form1());
        }
    }
}

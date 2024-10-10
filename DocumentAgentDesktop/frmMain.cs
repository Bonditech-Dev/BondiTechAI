using System;
using System.Windows.Forms;
//using DocumentAgentService;

namespace DocumentAgentDesktop
{
    public partial class frmMain : Form
    {

        DocumentAgentService.DocumentAgentService obj; 
        const string ServiceConfigFile = "documentagentservice.exe";
        public frmMain()
        {
           obj = new DocumentAgentService.DocumentAgentService();
            int NumberOfDaysPassed  = 0; // number of days using the trial

            InitializeComponent();
            obj.ReadConfigSettingsFromAppSettings(ServiceConfigFile);

            if (obj.objectDA.IsTrialVersion(ref NumberOfDaysPassed))
            {
                string message;
                string caption; string ContactMessage = "For information on license registration, please contact sales@casenamics.com.";
                caption = "Casenamics Barcode Router";
                MessageBoxButtons buttons = MessageBoxButtons.OK;
                DialogResult result;
                if (NumberOfDaysPassed > 30)
                {
                    message = "Your trial has expired. " + ContactMessage;
                    result = MessageBox.Show(message, caption, buttons,MessageBoxIcon.Information);
                    WriteMessage(message);
                }
                else
                {
                    WriteMessage("You are running a trial version. " + ContactMessage);
                    message = "You have " + (30 - NumberOfDaysPassed).ToString() + " days left on your trial. ";
                    result = MessageBox.Show(message + " " + ContactMessage, caption, buttons, MessageBoxIcon.Information);
                    WriteMessage(message);
                }
            }

        }

        private string GetDateTime()
        {
            return (DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString());
        }

        private void WriteMessage(string messageText)
        {
            txtMessage.Text += GetDateTime() + " - " + messageText + Environment.NewLine;
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            StartService();
        }

        void StartService()
        {
            // read from the documentagentservice.exe.config file instead of the default app config log.
            obj.ReadConfigSettingsFromAppSettings(ServiceConfigFile);
            obj.StartTimer();
            WriteMessage("Service started.");  
            WriteMessage("Looking for scanned documents in folder, " + obj.objectDA.Folder_Input + ".");
        }

        void StopService()
        {
            WriteMessage("Service stopped.");
            obj.StopTimer();
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            StopService();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            DialogResult ret = MessageBox.Show("Are you sure you want to exit?", ProductName, MessageBoxButtons.OKCancel);
            if (ret == DialogResult.OK)
            {
                obj.StopTimer();
                Close();
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox frm = new AboutBox();
            frm.ShowDialog();
        }

        private void startToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StartService();
        }

        private void stopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StopService();
        }

        private void settingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DocumentAgent.ConfigureSettings TestForm = new DocumentAgent.ConfigureSettings();
            TestForm.ConfigurationFile = Properties.Settings.Default.ConfigFile; //AppDomain.CurrentDomain.SetupInformation.ConfigurationFile;
            TestForm.LoadSettings();
            TestForm.ShowDialog();
        }

        private void viewLogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmViewLog frm = new frmViewLog();
            frm.MessageLine1 = @"Log below is for Barcode Reader.  Application logs may also  be found in the Windows Event Viewer.";
            frm.MessageLine2 = @"To view the log output from the Powershell script, open the Powershell script to find the location of the log. ";
            frm.FormTitle = "View Log";
            frm.LogFile = obj.GetLogFile();
            frm.ShowDialog();

        }


        private void helpToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            frmViewLog frm = new frmViewLog();
            frm.MessageLine1 = @"HELP for Barcode Reader";
            frm.MessageLine2 = @"";
            frm.FormTitle = "Barcode Reader Help";
            frm.LogFile = "Help.txt";
            frm.ShowDialog();
        }
        
    }
}

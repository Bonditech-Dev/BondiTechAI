using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

// TODO: Need to pass the configuraiton file to this somehow... Need to know for the Windows service.. 
// TODO: Allow them to configure the email notification here too ??

namespace DocumentAgent
{
    public partial class ConfigureSettings : Form
    {
        public string ConfigurationFile; // pass the configuration file..

        public ConfigureSettings()
        {
            InitializeComponent();

            PopulateListView();
            lblFolderInput.Text = "";
            //LoadSettings();
        }

        void PopulateListView()
        {
            string[,] ValueList = new string[,]     {
                                                    { "AZTEC","Aztec 2D barcode format"},
                                                    { "CODABAR","CODABAR 1D format"},
                                                    { "CODE_39","Code 39 1D format"},
                                                    { "CODE_93","Code 93 1D format"},
                                                    { "CODE_128","Code 128 1D format"},
                                                    { "DATA_MATRIX","Data Matrix 2D barcode format"},
                                                    { "EAN_8","EAN - 8 1D format."},
                                                    { "EAN_13","EAN - 13 1D format."},
                                                    { "ITF","ITF(Interleaved Two of Five) 1D format"},
                                                    { "MAXICODE","MaxiCode 2D barcode format"},
                                                    { "PDF_417","PDF417 format."},
                                                    { "QR_CODE","QR Code 2D barcode format"},
                                                    { "RSS_14","RSS 14"},
                                                    { "RSS_EXPANDED","RSS EXPANDED"},
                                                    { "UPC_A","UPC - A 1D format"},
                                                    { "UPC_E","UPC - E 1D format"},
                                                    { "UPC_EAN_EXTENSION","UPC / EAN extension format. Not a stand - alone format"},
                                                    { "MSI","MSI"},
                                                    { "PLESSEY","Plessey"},
                                                    { "ALL_1D"," UPC_A | UPC_E | EAN_13 | EAN_8 | CODABAR | CODE_39 | CODE_93 | CODE_128 | ITF | RSS_14 | RSS_EXPANDED without MSI (to many false - positives)"}
                                                };

            try
            {
                lvwBarcodeFormats.Columns.Clear();
                lvwBarcodeFormats.Columns.Add("Barcode Format", 100);
                lvwBarcodeFormats.Columns.Add("Description", 500);

                lvwBarcodeFormats.Items.Clear();

                for (int i = 0; i != ValueList.GetLength(0); i++)
                {
                    ListViewItem row = new ListViewItem(ValueList[i, 0]);
                    row.SubItems.Add(ValueList[i, 1]);
                    row.Tag = ValueList[i, 0];
                    lvwBarcodeFormats.Items.Add(row);
                }
            }
            catch (Exception exceptionDetails)
            {
                MessageBox.Show(exceptionDetails.Message, ProductName, MessageBoxButtons.OK);
            }

        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            SaveSettings();
        }

        public void LoadSettings()
        {

            try
            {
                // TODO: read these settings from the configuration file.. 
                // Make sure this config file exists in the DEBUG folder when debugging and filename matches.. 
                System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationFile);

                string InputFolder = config.AppSettings.Settings["Folder_Root"].Value;
                txtInputFolder.Text = InputFolder;
                UpdateInputLabel();

                string BarcodeFormats = config.AppSettings.Settings["BarcodeFormats"].Value;
                // split them by commas and check off the items in the list view...

                string[] BarcodeFormatsList = BarcodeFormats.Split(',');
                for (int i = 0; i <= BarcodeFormatsList.Length; i++)
                {

                    for (int k = 0; k != lvwBarcodeFormats.Items.Count; k++)
                    {
                        if (lvwBarcodeFormats.Items[k].Tag.ToString() == BarcodeFormatsList[i])
                        {
                            // this this item to checked... 
                            lvwBarcodeFormats.Items[k].Checked = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
               // MessageBox.Show(exceptionDetails.Message, ProductName, MessageBoxButtons.OK);
               // Ignore exception since setting may not have existed yet.. 
            }
        }

        private void SaveSettings()
        {
            try
            {
                string ScriptFolder = "";
                   
                if (txtInputFolder.Text.Trim().Length == 0)
                {
                    MessageBox.Show("Input Folder is required.", ProductName, MessageBoxButtons.OK);
                }
                else
                {
                    System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationFile);

                    config.AppSettings.Settings.Remove("Folder_Root");
                    config.AppSettings.Settings.Add("Folder_Root", txtInputFolder.Text);

                    // add default settings for rest of folder.. 
                    config.AppSettings.Settings.Remove("Folder_Input");
                    config.AppSettings.Settings.Add("Folder_Input", txtInputFolder.Text + @"\Input");

                    //9-7-2018, bpath, add temp folder..
                    config.AppSettings.Settings.Remove("Folder_Input_Temp");
                    config.AppSettings.Settings.Add("Folder_Input_Temp", txtInputFolder.Text + @"\Work\Temp");

                    config.AppSettings.Settings.Remove("Folder_Input_Processed");
                    config.AppSettings.Settings.Add("Folder_Input_Processed", txtInputFolder.Text + @"\Processed");

                    config.AppSettings.Settings.Remove("Folder_Input_Work");
                    config.AppSettings.Settings.Add("Folder_Input_Work", txtInputFolder.Text + @"\Work");

                    config.AppSettings.Settings.Remove("Folder_Output_Success");
                    config.AppSettings.Settings.Add("Folder_Output_Success", txtInputFolder.Text + @"\Success");


                    config.AppSettings.Settings.Remove("Folder_Output_Error");
                    config.AppSettings.Settings.Add("Folder_Output_Error", txtInputFolder.Text + @"\Error");


                    config.AppSettings.Settings.Remove("Number_Of_Threads");
                    config.AppSettings.Settings.Add("Number_Of_Threads", "1");


                    config.AppSettings.Settings.Remove("LogFileName");
                    config.AppSettings.Settings.Add("LogFileName", txtInputFolder.Text + @"\Service.log");

                    ScriptFolder = txtInputFolder.Text + @"\Scripts";
                    config.AppSettings.Settings.Remove("Folder_Scripts");
                    config.AppSettings.Settings.Add("Folder_Scripts", txtInputFolder.Text + @"\Scripts");


                    string SelectedBarCodes = "";
                    foreach (ListViewItem item in lvwBarcodeFormats.Items)
                    {
                        if (item.Checked)
                        {
                            SelectedBarCodes += SelectedBarCodes.Length > 0 ? "," : "";
                            SelectedBarCodes += item.SubItems[0].Text;
                        }
                    }
                    config.AppSettings.Settings.Remove("BarcodeFormats");
                    config.AppSettings.Settings.Add("BarcodeFormats", SelectedBarCodes);

                    // Save the configuration file.
                    config.Save(ConfigurationSaveMode.Modified, true);

                    // everything ok, disable Apply button now..
                    btnApply.Enabled = false;

                    Directory.CreateDirectory(ScriptFolder);
                    string MessageText = @"Please make sure your Powershell script file, processfiles.ps1, is located in location " + ScriptFolder + Environment.NewLine;
                    MessageText += @"Also, make sure any references to this location in the script are updated accordingly." + Environment.NewLine;
                    MessageText += @"Sample Powershell scripts are located in the application's Script folder (ex.c:\Program Files (x86)\Casenamics\Casenamics Barcode Router\Scripts).";
                    MessageBox.Show(MessageText , ProductName, MessageBoxButtons.OK);

                }
            }
            catch (Exception exceptionDetails)
            {
                MessageBox.Show(exceptionDetails.Message, ProductName, MessageBoxButtons.OK);
            }


        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {

            try
            {
                FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
                DialogResult result = folderBrowserDialog1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    txtInputFolder.Text = folderBrowserDialog1.SelectedPath;
                }
            }
            catch (Exception exceptionDetails)
            {
                MessageBox.Show(exceptionDetails.Message, ProductName, MessageBoxButtons.OK);
            }


        }

        private void txtInputFolder_TextChanged(object sender, EventArgs e)
        {
            btnApply.Enabled = true;
            UpdateInputLabel();
        }

        void UpdateInputLabel()
        {
            lblFolderInput.Text = txtInputFolder.Text + @"\input";
        }
        private void btnOK_Click(object sender, EventArgs e)
        {
            SaveSettings();
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void lvwBarcodeFormats_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            btnApply.Enabled = true;
        }

        private void btnSubmitIssue_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.casenamics.com/Doc/SubmitIssue");
        }
    }
}

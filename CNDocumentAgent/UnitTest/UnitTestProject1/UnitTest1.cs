using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Casenamics;
using System.Windows.Forms;
using System.IO;
using System.Linq;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {

        [TestMethod]
        public void TestMethod1()
        {

            StartAgent();
            //RunAgent();
        }


        [TestMethod]
        public void GetFileCount()
        {
            int cnt = 0;
            var FileList = Directory.EnumerateFiles(@"c:\temp\testing", "*.daxml", SearchOption.TopDirectoryOnly);
            foreach (string currentFile in FileList)
            {
                DateTime FileDate = File.GetCreationTime(currentFile);
                if (FileDate.Date == System.DateTime.Now.Date) cnt++;
            }
        }

[TestMethod]
        public void TestRead()
        {
            Casenamics.DocumentAgent objectDA = new Casenamics.DocumentAgent();

            objectDA.BarcodeFormats = "CODE_39";

            // objectDA.TestFile(@"C:\TEMP\SCAN\labels.tif");
            //objectDA.TestFile(@"C:\TEMP\SCAN\image0014.tif");

        }

        [TestMethod]
        public void TestPDFSharpImage()
        {
            Casenamics.DocumentAgent objectDA = new Casenamics.DocumentAgent();
            //           objectDA.ExtractImageFromPDF(@"C:\Workbnch\Casenamics\Source\CNDocumentAgent\Scanning\Testing\Sample Documents\132132_Office2010.pdf"); 
            objectDA.ExtractImageFromPDF(@"C:\VPCShared\Casenamics\Scanning\Input\3.pdf"); 
            // C:\VPCShared\Casenamics\Scanning\Input
        }

        void StartAgent()
        {
            Casenamics.DocumentAgent objectDA = new Casenamics.DocumentAgent();
            objectDA.BarcodeFormats = "CODE_39";

            objectDA.Folder_Input = @"C:\VPCShared\Casenamics\Scanning\Input";
            objectDA.Folder_Input_Processed = @"C:\VPCShared\Casenamics\Scanning\Input\Processed";
            objectDA.Folder_Input_Work = @"C:\VPCShared\Casenamics\Scanning\Input_Work";
            objectDA.Folder_Output_Error = @"C:\VPCShared\Casenamics\Scanning\Output_Error";
            objectDA.Folder_Output_Success = @"C:\VPCShared\Casenamics\Scanning\Output_Success";
            objectDA.Folder_Scripts = @"C:\VPCShared\Casenamics\Scanning\Scripts";
            objectDA.LogFileName = @"C:\VPCShared\Casenamics\Scanning\Scripts\Script.log";
            //bjectDA.Number_Of_Threads = -1;
            objectDA.TimerInterval = 30;
            objectDA.VerboseLogging = true;

            objectDA.Number_Of_Threads = 1;

            //objectDA.ExecuteSynchronousScript();
            objectDA.StartAgent(); // starts it up.. Should put this call on a timer if we want the polling to be on... 
        }

        void RunAgent()
        {
            Casenamics.DocumentAgent objectDA = new Casenamics.DocumentAgent();

            objectDA.Folder_Input = @"C:\Workbnch\Casenamics\Source\CNDocumentAgent\Scanning\Input";
            objectDA.Folder_Input_Processed = @"C:\Workbnch\Casenamics\Source\CNDocumentAgent\Scanning\Input\Processed";
            objectDA.Folder_Input_Work = @"C:\Workbnch\Casenamics\Source\CNDocumentAgent\Scanning\Input_Work";
            objectDA.Folder_Output_Error = @"C:\Workbnch\Casenamics\Source\CNDocumentAgent\Scanning\Output_Error";
            objectDA.Folder_Output_Success = @"C:\Workbnch\Casenamics\Source\CNDocumentAgent\Scanning\Output_Success";
            objectDA.Folder_Scripts = @"C:\Workbnch\Casenamics\Source\CNDocumentAgent\Scanning\Scripts";
            objectDA.LogFileName = @"C:\Workbnch\Casenamics\Source\CNDocumentAgent\Scanning\Scripts\Script.log";
            objectDA.Number_Of_Threads = 1;

            //objectDA.ExecuteSynchronousScript();
            objectDA.RunScript(); // starts it up.. Should put this call on a timer if we want the polling to be on... 
        }



       
        public void TestRunScript()
        {
            Casenamics.DocumentAgent objectDA = new Casenamics.DocumentAgent();

            objectDA.Folder_Input = @"C:\Workbnch\Casenamics\Source\CNDocumentAgent\Scanning\Input";
            objectDA.Folder_Input_Work = @"C:\Workbnch\Casenamics\Source\CNDocumentAgent\Scanning\Input_Work";
            objectDA.Folder_Output_Error = @"C:\Workbnch\Casenamics\Source\CNDocumentAgent\Scanning\Output_Error";
            objectDA.Folder_Output_Success = @"C:\Workbnch\Casenamics\Source\CNDocumentAgent\Scanning\Output_Success";
            objectDA.Folder_Scripts = @"C:\Workbnch\Casenamics\Source\CNDocumentAgent\Scanning\Scripts";
            objectDA.LogFileName = @"C:\Workbnch\Casenamics\Source\CNDocumentAgent\Scanning\Scripts\Script.log";
            objectDA.Number_Of_Threads = 1;

            //objectDA.ExecuteSynchronousScript();
            //objectDA.StartAgent(); // starts it up.. Should put this call on a timer if we want the polling to be on... 
            objectDA.RunScript();

        }

       
        public void TestConfigurationSettingsForm()
        {
            DocumentAgent.ConfigureSettings FormSettings = new DocumentAgent.ConfigureSettings();
            FormSettings.ShowDialog();
          
        }


     }



}

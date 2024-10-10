using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocumentAgentDesktop
{
    public partial class frmViewLog : Form
    {
        public string LogFile = "";
        public string MessageLine1 = "";
        public string MessageLine2 = "";
        public string FormTitle = "";

        public frmViewLog()
        {
            InitializeComponent();
        }

        void LoadLog()
        {
            try
            {
                txtLog.Text = System.IO.File.ReadAllText(LogFile);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to open log file, " + LogFile + "." + Environment.NewLine + "Error: " + ex.Message);
            }
        }
        private void frmViewLog_Load(object sender, EventArgs e)
        {

            LoadLog();
            lblMessage.Text = MessageLine1;
            lblMessage2.Text = MessageLine2;
            this.Text = FormTitle;
           // frmViewLog.Titl

            //myfilestream = New FileStream(fileName, FileMode.Open, FileAccess.Read)
            //Dim count As Integer ' = 1002 ' fixed width of bulk load files for ovs vendor... 

            //count = System.Convert.ToInt16(My.Settings.IDLRecordLength)
            //If count <= 0 Then count = IDL_RECORD_LENGTH ' then default it if there is no setting from configuration file.. 
            //Dim buffer(count - 1) As Byte
            //Dim bufferString As String = ""
            //frm.Show()
            //LogFile = fileName & ".log"
            //frm.LogFile = LogFile
            //frm.lblMessage.Text = "Processing file. To see log, open " & LogFile & "."
            //LogText = "Process started " & Format(Now(), "General Date") & vbCrLf _
            //& "Opening File " & fileName & vbCrLf

            //WriteLogToFile(LogFile, "---------------------------------------------------------------------------------")
            //frm.txtOutput.AppendText(LogText)
            //WriteLogToFile(LogFile, LogText)
            //myfilestream.Read(buffer, 0, IDL_RECORD_LENGTH) ' skip first two header record lines... 
            //myfilestream.Read(buffer, 0, IDL_RECORD_LENGTH)
            //While count <> 0
            //    ' Get values from each line.. 
            //    Contents = ""
            //    count = myfilestream.Read(buffer, 0, count)
            //    CopyBytesToString(buffer, Contents)
            //    If count > 10 Then
            //        LineCount += 1
            //        ' skip header if linecount = 1
            //        'If LineCount <> 1 AndAlso GetRecordDetails(Contents, VendorRecord) Then
            //        If GetIDLRecordDetails(Contents, IDLRecord) Then
            //            MessageText = "Doc Number = " & IDLRecord.IDL_DOCUMENT_NUMBER
            //            frm.txtOutput.AppendText("Record#" & LineCount.ToString & ", Processing " & MessageText & vbCrLf)
            //            Application.DoEvents()
            //            ' will update if existing.. 
            //            InsertIDLInfo(IDLRecord, LogFile, ErrorOccurred)
            //        End If
            //    End If
            //End
        }
    }
}

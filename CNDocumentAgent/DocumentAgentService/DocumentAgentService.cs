using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.ServiceProcess;
using System.Threading;

// if installing this web service, make sure you compile in release mode, or it will do the debugonstart instead!!!!!!! 
namespace DocumentAgentService
{
    public partial class DocumentAgentService : ServiceBase
    {

        public Casenamics.DocumentAgent objectDA = new Casenamics.DocumentAgent();
        private Thread thread1, thread2;


        public DocumentAgentService()
        {
            try
            {
                InitializeComponent();
                ReadConfigSettingsFromAppSettings();
            }
            catch (ConfigurationErrorsException configEx)
            {
                WriteEventLog("Configuration error in DocumentAgentService: " + configEx.Message, EventLogEntryType.Error);
            }
            catch (Exception ex)
            {
                WriteEventLog("Unexpected error in DocumentAgentService: " + ex.Message, EventLogEntryType.Error);
            }
        }


        public void DebugOnStart()
        {
            //WriteEventLog("Hello");

            try
            {
                //ProcessExistingFiles();
                OnStart(null);

            }
            catch (Exception ex)
            {
                WriteEventLog("DebugOnStart:" + ex.Message, EventLogEntryType.Error);
            }

        }

        protected override void OnStart(string[] args)
        {
            try
            {

                // launch thread for processing existing files.. 
                //WriteEventLog("Service Started. thread2", EventLogEntryType.Information);
                thread2 = new Thread(new ThreadStart(ProcessExistingFiles));
                thread2.IsBackground = true;
                thread2.Start();
                //WriteEventLog("Service Started. thread2 - started", EventLogEntryType.Information);

                // TODO: This thread does not return when service is started... 

                // launch another thread but have it wait until thread2 is done. MonitorForNewFiles();
                //WriteEventLog("Service Started. thread1", EventLogEntryType.Information);
                thread1 = new Thread(new ThreadStart(MonitorForNewFiles));
                thread1.IsBackground = true;
                thread1.Start();
                //WriteEventLog("Service Started. thread1 - started", EventLogEntryType.Information);

                string AppVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();
                WriteEventLog("Service Started. Version, " + AppVersion, EventLogEntryType.Information);

            }
            catch (Exception ex)
            {
                WriteEventLog("OnStart:" + ex.Message, EventLogEntryType.Error);
            }

            //ProcessExistingFiles();
            //MonitorForNewFiles();

            // put out a temporary file to trigger the FileSystemWatcher to process existing files... 
            // string TempFileName = objectDA.Folder_Input + @"\temp.txt";
            // System.IO.File.Create(TempFileName);
            // System.IO.File.Delete(TempFileName);

            //thread1 = new Thread(new ThreadStart(MonitorForNewFiles));
            //thread1.IsBackground = true;
            //thread1.Start();

        }

        protected override void OnStop()
        {
            try
            {
                if (objectDA.Number_Of_Threads==-1)
                {
                    StopTimer();
                    WriteEventLog("Timer stopped", EventLogEntryType.Information);
                }
  
                // Clean up?
                thread1.Abort();
                thread2.Abort();
                WriteEventLog("Service Stopped.", EventLogEntryType.Information);

            }
            catch (Exception ex)
            {
                WriteEventLog("OnStop Clicked multiple" + ex.Message, EventLogEntryType.Error);
                WriteEventLog("OnStop" + ex.Message, EventLogEntryType.Error);
            }

        }

        private void MonitorForNewFiles()
        {
            try
            {
                if (objectDA.Number_Of_Threads==-1)
                {
                    // User timer instead..
                    StartTimer();
                    WriteEventLog("Timer set to monitor new files in folder, " + objectDA.Folder_Input, EventLogEntryType.Information);
                    WriteEventLog("Timer started", EventLogEntryType.Information);
                }
                else
                {
                    thread2.Join(); // wait until thread2 is done before we execute this thread.. 
                    WriteEventLog("Monitoring New Files in folder, " + objectDA.Folder_Input, EventLogEntryType.Information);
                    FileSystemWatcher Watcher = new FileSystemWatcher(objectDA.Folder_Input);
                    Watcher.EnableRaisingEvents = true;
                    Watcher.Created += new FileSystemEventHandler(FileCreated);
                    Watcher.Error += new ErrorEventHandler(OnError);
                    Watcher.Changed += new FileSystemEventHandler(FileChanged);
                    Watcher.Renamed += new RenamedEventHandler(FileRenamed);
                    //Watcher.Deleted += new FileSystemEventHandler(FileDeleted);
                    WriteEventLog("Event Handlers enabled.", EventLogEntryType.Information);
                }

            }
            catch (Exception ex)
            {
                WriteEventLog("MonitorForNewFiles:" + ex.Message, EventLogEntryType.Error);
            }
        }

        private  void OnError(Object source, ErrorEventArgs e)
        {
            WriteEventLog(e.GetException().Message, EventLogEntryType.Error);

        }

        private void FileCreated(Object sender, FileSystemEventArgs e)
        {
            WriteEventLog("Detected FileCreated.", EventLogEntryType.Information);
            try
            {
                // e.Name and e.FullPath are the files.. 
                ProcessFiles();

            }
            catch (Exception ex)
            {
                WriteEventLog("FileCreated:" + ex.Message, EventLogEntryType.Error);
            }

        }


        private void FileChanged(Object sender, FileSystemEventArgs e)
        {
            WriteEventLog("Detected FileChanged.", EventLogEntryType.Information);
        }

        private void FileRenamed(Object sender, FileSystemEventArgs e)
        {
            WriteEventLog("Detected FileRenamed.", EventLogEntryType.Information);
        }

        void ProcessExistingFiles()
        {
            WriteEventLog("Processing Existing Files.", EventLogEntryType.Information);
            try
            {
                //string[] FileList = Directory.GetFiles(objectDA.Folder_Input);
                if (FilesExistInInputFolder())
                {
                    ProcessFiles();
                }

            }
            catch (Exception ex)
            {
                WriteEventLog("ProcessExistingFiles:" + ex.Message, EventLogEntryType.Error);
            }

        }

        bool FilesExistInInputFolder()
        {
            bool RetVal = false;
            string[] FileList = Directory.GetFiles(objectDA.Folder_Input);
            if (FileList.Length > 0)
            {
                // Define supported file extensions
                var supportedExtensions = new HashSet<string> { ".tiff", ".bmp", ".png", ".gif", ".pdf" };

                // Check if any file has a supported extension
                foreach (var file in FileList)
                {
                    string fileExtension = Path.GetExtension(file).ToLower();
                    if (supportedExtensions.Contains(fileExtension))
                    {
                        RetVal = true;
                        break;
                    }
                }
            }

            return RetVal;
        }


        void ProcessFiles()
        {
            WriteEventLog("Process Files.", EventLogEntryType.Information);

            try
            {
                // keep processing files until all files are processed.. 
                while (FilesExistInInputFolder())
                {
                    objectDA.StartAgent();
                    objectDA.RunScript();
                }

            }
            catch (Exception ex)
            {
                WriteEventLog("ProcessFiles:" + ex.Message, EventLogEntryType.Error);
            }
        }

        // reads based on AppSettings instead of ApplicationSettings..
        public void ReadConfigSettingsFromAppSettings(string configFile = "")
        {
            try
            {
                Configuration config = configFile.Length > 0
                    ? ConfigurationManager.OpenExeConfiguration(configFile)
                    : ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

                var settings = config.AppSettings.Settings;

                objectDA.Folder_Input = settings["Folder_Input"].Value;
                objectDA.Folder_Input_Temp = settings["Folder_Input_Temp"].Value;
                objectDA.Folder_Input_Processed = settings["Folder_Input_Processed"].Value;
                objectDA.Folder_Input_Work = settings["Folder_Input_Work"].Value;
                objectDA.Folder_Output_Error = settings["Folder_Output_Error"].Value;
                objectDA.Folder_Output_Success = settings["Folder_Output_Success"].Value;
                objectDA.Folder_Scripts = settings["Folder_Scripts"].Value;
                objectDA.LogFileName = settings["LogFileName"].Value;
                objectDA.BarcodeFormats = settings["BarcodeFormats"].Value;
                objectDA.Number_Of_Threads = Convert.ToInt32(settings["Number_Of_Threads"].Value);
                objectDA.TimerInterval = Convert.ToInt32(settings["TimerInterval"].Value);
                objectDA.VerboseLogging = settings["VerboseLogging"].Value == "1";

                // if blank, then default to input folder
                if (string.IsNullOrEmpty(objectDA.Folder_Input_Temp))
                {
                    objectDA.Folder_Input_Temp = objectDA.Folder_Input;
                }

                // create folders if they don't exist
                objectDA.CreateFolders();

                // set the license and the script file name
                objectDA.LicenseFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cndalicense.lic");

                // get the trial file
                objectDA.TrialFile = Path.Combine(objectDA.Folder_Input_Processed, "cnda.dat");
            }
            catch (ConfigurationErrorsException configEx)
            {
                WriteEventLog("Configuration error in ReadConfigSettingsFromAppSettings: " + configEx.Message, EventLogEntryType.Error);
            }
            catch (Exception ex)
            {
                WriteEventLog("Unexpected error in ReadConfigSettingsFromAppSettings: " + ex.Message, EventLogEntryType.Error);
            }
        }


        // No longer used, using AppSettings instead.. With AppSettings, we are able to modify from a different program. 
        //void ReadConfigSettings()
        //{
        //    objectDA.Folder_Input = Properties.Settings.Default.Folder_Input; 
        //    objectDA.Folder_Input_Processed = Properties.Settings.Default.Folder_Input_Processed; 
        //    objectDA.Folder_Input_Work = Properties.Settings.Default.Folder_Input_Work; 
        //    objectDA.Folder_Output_Error = Properties.Settings.Default.Folder_Output_Error; 
        //    objectDA.Folder_Output_Success = Properties.Settings.Default.Folder_Output_Success; 
        //    objectDA.Folder_Scripts = Properties.Settings.Default.Folder_Scripts; 
        //    objectDA.LogFileName = Properties.Settings.Default.LogFileName; 
        //    objectDA.BarcodeFormats = Properties.Settings.Default.BarcodeFormats; 
        //    objectDA.Number_Of_Threads = 1;
        //}

        // if you get an error with the below statement, then make sure you run as administrator. It doesn't have writes to read the event log.. 
        void WriteEventLog(string messageText, EventLogEntryType logType)
        {
            // must run or debug as admin if you want to write to event view... 
            try
            {
                const string LogName = "Casenamics Document Agent Service";

                if (!EventLog.SourceExists(LogName))
                {
                    EventLog.CreateEventSource(LogName, LogName);
                }

                using (EventLog appLog = new EventLog { Source = LogName })
                {
                    appLog.WriteEntry(messageText, logType);
                }
            }
            catch (Exception ex)
            {
                // Log the exception to a fallback mechanism if necessary
                // For now, we are just swallowing the exception to avoid crashing the service
            }
        }


        public string GetLogFile()
        {
            return !string.IsNullOrEmpty(objectDA.LogFileName)
                ? objectDA.LogFileName
                : Path.Combine(objectDA.Folder_Scripts, "service.log");
        }


        // Use Time instead..
        private System.Threading.Timer FileMonitorTimer;


        // Does not work on multi threads right now. Timers will create new thread.
        // Currently, we just disable the timer so it doesn't create anther thread.
        // to make multi threaded, you need to make timer friendly with thread using AutoResetEvent and 
        // passing it to timercallback fcn. 
        public void StartTimer()
        {
            try
            {
                if (FileMonitorTimer == null)
                {
                    int timerIntervalMillis = objectDA.TimerInterval;
                    WriteEventLog($"StartTimer - set every {timerIntervalMillis} milliseconds", EventLogEntryType.Information);

                    FileMonitorTimer = new System.Threading.Timer(
                        Timer_Tick,
                        null,
                        0,
                        timerIntervalMillis
                    );
                }
            }
            catch (Exception ex)
            {
                WriteEventLog("StartTimer: " + ex.Message, EventLogEntryType.Error);
            }
        }

        public void StopTimer()
        {
            try
            {
                if (FileMonitorTimer != null)
                {
                    WriteEventLog("StopTimer", EventLogEntryType.Information);
                    FileMonitorTimer.Change(Timeout.Infinite, Timeout.Infinite);
                    FileMonitorTimer.Dispose();
                    FileMonitorTimer = null;
                }
            }
            catch (ObjectDisposedException)
            {
                WriteEventLog("StopTimer: Timer already disposed.", EventLogEntryType.Warning);
            }
            catch (Exception ex)
            {
                WriteEventLog("StopTimer: " + ex.Message, EventLogEntryType.Error);
            }
        }


        private void Timer_Tick(object state)
        {
            try
            {
                if (objectDA.VerboseLogging) WriteEventLog("TimerTick", EventLogEntryType.Information);

                // stop timer so it doesn't call this tick while it's still processing.. 
                FileMonitorTimer.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite);
                //StopTimer();

                ProcessFiles();

                //StartTimer();
                FileMonitorTimer.Change(objectDA.TimerInterval, objectDA.TimerInterval);

                //// restart timer.. 
                //int timerIntervalSecs = 10; // TODO: read from config. 
                //TimeSpan tsInterval = new TimeSpan(0, 0, timerIntervalSecs);
                //FileMonitorTimer.Change(tsInterval, tsInterval);


            }
            catch (Exception ex)
            {
                WriteEventLog("Timer_Tick: " + ex.Message, EventLogEntryType.Error);
            }
        }


    }
}

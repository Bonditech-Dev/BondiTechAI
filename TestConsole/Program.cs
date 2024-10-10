using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


// just test the filesystemwatch object..

namespace TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hit any key to exit");
            MonitorForNewFiles();
            Console.ReadLine();
        }


        private static void MonitorForNewFiles()
        {
            try
            {
                string FolderName = Properties.Settings.Default.Folder;  //ConfigurationManager.AppSettings["Folder"];
                FileSystemWatcher Watcher = new FileSystemWatcher(FolderName);
                Watcher.EnableRaisingEvents = true;
                Watcher.Created += new FileSystemEventHandler(FileCreated);
                Watcher.Error += new ErrorEventHandler(OnError);
               // Watcher.NotifyFilter = NotifyFilters.FileName;
                //Watcher.InternalBufferSize = 25 * 4096;
                Console.WriteLine("MonitorForNewFiles, Folder=" + FolderName);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private static void OnError(Object source, ErrorEventArgs e)
        {
            Console.WriteLine("OnError - " + e.GetException().Message);
        }

        private static void FileCreated(Object sender, FileSystemEventArgs e)
        {
            Console.WriteLine("File Created");

        }


    }
}

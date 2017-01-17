using System;
using System.ServiceProcess;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace Folder_Monitor
{
    public partial class TranscriptMonitor : ServiceBase
    {
        public TranscriptMonitor()
        {
            // Basic service settings
            this.ServiceName = "FolderMonitor";
            this.CanStop = true;
            this.CanPauseAndContinue = false;
            this.AutoLog = true;
        }

        public static void WriteServiceLog(String LogMessage)
        {
            // Function to write to the service log at the below path
            StreamWriter ServiceLog = new System.IO.StreamWriter(@"E:\ServiceLog.txt", true);
            String CurrentTime = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
            ServiceLog.WriteLine("[" + CurrentTime + "] " + LogMessage); // Outputs the file
            ServiceLog.Close(); // Closes the file
        }

        public static void Monitor()
        {
            // The core of the program. this will keep an eye on the folder itself for any updates/changes
            WriteServiceLog("Starting Transcript Monitor");
            FileSystemWatcher FolderMonitor = new System.IO.FileSystemWatcher();
            FolderMonitor.Path = "E:\\FTP\\Transcripts\\OCME\\Dictators";
            FolderMonitor.IncludeSubdirectories = true;

            FolderMonitor.NotifyFilter = NotifyFilters.LastAccess |
                NotifyFilters.LastWrite |
                NotifyFilters.FileName |
                NotifyFilters.DirectoryName;

            FolderMonitor.Filter = "*.doc*"; // File extensions (this one searches for .doc and .docx

            FolderMonitor.Created += new FileSystemEventHandler(OnChanged);
            FolderMonitor.EnableRaisingEvents = true; // Enables the event monitor

            WriteServiceLog("File Watching started");
        }
        public static void OnChanged(object source, FileSystemEventArgs e)
        {
            // Specify what is done when a file is changed.
            // e.Name, e.FullPath, e.ChangeType);
            WriteServiceLog("File uploaded - " + e.FullPath);

            string command = @"/c powershell -ExecutionPolicy Unrestricted E:\Scripts\TranscriptMonitor-v3.ps1 " + e.FullPath;
            WriteServiceLog("Executing command: " + command);
            System.Diagnostics.Process.Start("cmd.exe", command);
            WriteServiceLog("Finished invoking script");
        }

        protected override void OnStart(string[] args)
        {
            WriteServiceLog("Service is starting");
            //InitializeComponent();
            System.ServiceProcess.ServiceBase.Run(new TranscriptMonitor());
            Monitor();
        }

        protected override void OnStop()
        {
            WriteServiceLog("Service is stopping");
            GC.Collect();
        }

        protected static void RunScript (string PSScript, string FileUploaded)
        {
            WriteServiceLog("PS Command: " + PSScript);
            WriteServiceLog("PS Argument: " + FileUploaded);
            // Call the InitailSessionState.CreateDefault method to create
            // an empty InitialSessionState object, and then add the 
            // elements that will be available when the runspace is opened.
            InitialSessionState iss = InitialSessionState.CreateDefault();

            WriteServiceLog("Created InitialSessionState");
            // Call the RunspaceFactory.CreateRunspace(InitialSessionState) 
            // method to create the runspace where the pipeline is run.
            Runspace rs = RunspaceFactory.CreateRunspace(iss);
            rs.Open();

            WriteServiceLog("Opening Runspace");
            
            // Call the PowerShell.Create() method to create the PowerShell 
            // object,and then specify the runspace and commands to the pipeline. 
            // and  create the command pipeline.
            PowerShell ps = PowerShell.Create();
            ps.Runspace = rs;
            ps.AddCommand(PSScript);
            ps.AddArgument(FileUploaded);

            WriteServiceLog("PS Command: " + PSScript);
            WriteServiceLog("PS Argument: " + FileUploaded);
            // Close the runspace to free resources.
            rs.Close();
        }

    }
}

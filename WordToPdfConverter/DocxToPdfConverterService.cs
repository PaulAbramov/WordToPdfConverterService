using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Timers;
using Microsoft.Office.Interop.Word;

namespace WordToPdfConverter
{
    public partial class DocxToPdfConverterService : ServiceBase
    {
        
        private int eventId = 1;

        private readonly string pathToFolder;

        public DocxToPdfConverterService(string[] _arguments)
        {
            InitializeComponent();

            eventLog1 = new EventLog();
            if (!EventLog.SourceExists("ConverterSource"))
            {
                EventLog.CreateEventSource(
                    "ConverterSource", "DocxToPdfConverterLog");
            }
            eventLog1.Source = "ConverterSource";
            eventLog1.Log = "DocxToPdfConverterLog";

            if (_arguments.Length > 1)
            {
                eventLog1.WriteEntry("too many arguments", EventLogEntryType.Error);
                return;
            }

            foreach (var argument in _arguments)
            {
                eventLog1.WriteEntry($"{argument}", EventLogEntryType.Warning);

            }

            pathToFolder = _arguments[0];
        }

        protected override void OnStart(string[] _args)
        {
            eventLog1.WriteEntry("OnStart");

            // Update the service state to Start Pending.
            ServiceStatus serviceStatus = new ServiceStatus
            {
                dwCurrentState = ServiceState.SERVICE_START_PENDING,
                dwWaitHint = 100000
            };
            SetServiceStatus(ServiceHandle, ref serviceStatus);

            Timer timer = new Timer();
            timer.Interval = 10000; // 60 seconds
            timer.Elapsed += OnTimer;
            timer.Start();

            // Update the service state to Running.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(ServiceHandle, ref serviceStatus);
        }

        private void OnTimer(object _sender, ElapsedEventArgs _args)
        {
            eventLog1.WriteEntry($"checking {pathToFolder}", EventLogEntryType.Information, eventId++);
            var docxFilesInFolder = new DirectoryInfo(pathToFolder).GetFiles();

            if (docxFilesInFolder.Any(x => x.Extension.Equals(".docx")))
            {
                eventLog1.WriteEntry($"there are {docxFilesInFolder.Count(x => x.Extension.Equals(".docx"))} docx files", EventLogEntryType.Information, eventId++);

                // Create an instance of Word.exe
                var wordApplication = new Application
                {
                    // Make this instance of word invisible
                    Visible = false,
                    ScreenUpdating = false,
                    DisplayAlerts = WdAlertLevel.wdAlertsNone,
                    Options =
                    {
                        SavePropertiesPrompt = false,
                        SaveNormalPrompt = false
                    }
                };

                object readOnly = false;
                object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                object fileFormat = WdSaveFormat.wdFormatPDF;
                object oMissing = System.Reflection.Missing.Value;
                foreach (var fileInfo in docxFilesInFolder)
                {
                    try
                    {
                        eventLog1.WriteEntry($"checking {fileInfo.FullName}", EventLogEntryType.Information, eventId++);

                        Object pathToFile = (Object)fileInfo.FullName;
                        object outputFileName = fileInfo.FullName.Replace(".docx", ".pdf");

                        if (File.Exists((string)outputFileName))
                        {
                            eventLog1.WriteEntry($"{(string)outputFileName} does exist already", EventLogEntryType.Information, eventId++);
                            continue;
                        }

                        // Load a document into our instance of word.exe
                        var document = wordApplication.Documents.Open(ref pathToFile, ref oMissing, ref readOnly, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                        eventLog1.WriteEntry($"changed path to {fileInfo.FullName.Replace(".docx", ".pdf")}", EventLogEntryType.Information, eventId++);

                        // Save document into PDF Format
                        document.SaveAs(ref outputFileName,
                            ref fileFormat, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                        // Close the Word document, but leave the Word application open.
                        // doc has to be cast to type _Document so that it will find the
                        // correct Close method.                
                        ((Document)document).Close(ref saveChanges, ref oMissing, ref oMissing);
                        document = null;
                    }
                    catch (Exception e)
                    {
                        eventLog1.WriteEntry($"{e.StackTrace}", EventLogEntryType.Information, eventId++);
                        throw;
                    }
                }

                // word has to be cast to type _Application so that it will find
                // the correct Quit method.
                ((_Application)wordApplication).Quit(ref oMissing, ref oMissing, ref oMissing);
                wordApplication = null;
            }
        }

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(System.IntPtr handle, ref ServiceStatus serviceStatus);

        public enum ServiceState
        {
            SERVICE_STOPPED = 0x00000001,
            SERVICE_START_PENDING = 0x00000002,
            SERVICE_STOP_PENDING = 0x00000003,
            SERVICE_RUNNING = 0x00000004,
            SERVICE_CONTINUE_PENDING = 0x00000005,
            SERVICE_PAUSE_PENDING = 0x00000006,
            SERVICE_PAUSED = 0x00000007,
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct ServiceStatus
        {
            public int dwServiceType;
            public ServiceState dwCurrentState;
            public int dwControlsAccepted;
            public int dwWin32ExitCode;
            public int dwServiceSpecificExitCode;
            public int dwCheckPoint;
            public int dwWaitHint;
        };
    }
}

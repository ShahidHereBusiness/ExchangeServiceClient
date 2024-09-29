using System;
using System.Diagnostics;
using System.IO;

namespace EWS.Includes
{

    public class HandleAppLog: IDisposable
    {
        private DateTime timeTrace;
        private string timeLog = string.Empty;
        private string timeStr = string.Empty;
        private string serverVariable = string.Empty;
        public HandleAppLog(string passServerVariable = "_")
        {
            timeTrace = DateTime.Now;
            timeLog = $"{DateTime.Now.ToString("yyyyMMddHHmmssfffffff")}{new Random(Guid.NewGuid().GetHashCode()).Next()}";                
            timeStr = DateTime.Now.ToString("yyyyMMddHH");
            if (!Validation.FormatError(passServerVariable))
                serverVariable = passServerVariable.Replace(":", "_");
        }
        public int FileSystemLog(string path, string MethodName, string direction, string logMsg)
        {
            try
            {
                if (!path.Substring(path.Length - 1).Contains("\\"))
                {
                    return 1;
                }

                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                // Compute the difference
                TimeSpan difference = DateTime.Now - timeTrace;
                string path2 = path + timeStr + "_" + serverVariable + ".txt";
                StreamWriter streamWriter = (File.Exists(path2) ? File.AppendText(path2) : File.CreateText(path2));
                streamWriter.WriteLine("TID:" + timeLog + "|TimeSpan:" + difference + "|" + MethodName + "|" + direction + "|" + logMsg);
                streamWriter.Close();
            }
            catch (Exception ex)
            {
                if (!EventViewerAppLog(MethodName + "\r\n" + logMsg + "\r\n" + ex.Message, "Application", 5175, 101))
                {
                    return 2;
                }

                return 3;
            }

            return 0;
        }
        public bool EventViewerAppLog(string logMsg, string source = "Application", int eventId = 5175, short category = 101)
        {
            try
            {
                if (!EventLog.SourceExists(source))
                {
                    return false;
                }

                EventLog eventLog = new EventLog(source);
                eventLog.Source = source;
                eventLog.WriteEntry(logMsg, EventLogEntryType.Information, eventId, category);
            }
            catch (Exception ex)
            {
                if (ex.Message.Length > 0)
                {
                    return false;
                }
            }
            return true;
        }

        public void Dispose()
        {
            return;
        }

        ~HandleAppLog()
        {
            timeLog = null;
            timeStr = null;
            serverVariable = null;
            GC.Collect();
        }
    }
}
using System;
using System.IO;

namespace StmMailDaemon.Models
{
    public class LogWriter
    {
        public static void WriteToLog(string message, bool timeStamp = true)
        {
            try
            {
                using var writer = new StreamWriter(GlobalVariables.LogPath, true);

                writer.WriteLine($"{(timeStamp ? DateTime.Now.ToString() + ": " : string.Empty)}{message}");
            }
            catch { }
        }
    }

}

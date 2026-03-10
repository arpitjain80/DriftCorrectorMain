using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentProcessor
{
    public class CustomLogger
    {
        string _logpath="";
        public CustomLogger(string LogPath)
        {
            _logpath = LogPath;
        }

        public void Log(string message)
        {
            if(_logpath == "")
            {
                return;
            }
            try
            {
                //string logPath = @"c:\Personal\Project\DriftCorrector\Files\Logs\logs.txt";
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(_logpath));
                System.IO.File.AppendAllText(_logpath, $"[{System.DateTime.Now:HH:mm:ss.fff}] {message}\r\n");
            }
            catch (System.Exception) { /* ignore */ }
        }

        public void Clear()
        {
            if (_logpath == "")
            {
                return;
            }
            try
            {
                //string logPath = @"c:\Personal\Project\DriftCorrector\Files\Logs\logs.txt";
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(_logpath));
                System.IO.File.WriteAllText(_logpath, "");
            }
            catch (System.Exception) { /* ignore */ }
        }
    }
}

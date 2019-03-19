using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ProjectX.Loger
{
    public class SystemLoger
    {
        private readonly string pathLog;

        public SystemLoger(string path) {
            pathLog = path;
            try
            {
                if (new FileInfo(path).Length > 42E+6)
                    throw new IOException();

            }
            catch (IOException) {
                File.Create(path).Close();
            }
            WriteLog("Сессия открыта");
        }

        public void WriteLog(string message) {
            lock (this) {
                
                using (StreamWriter sw = new StreamWriter(pathLog, true)) {
                    sw.WriteLine(DateTime.Now + ":\t"+message);
                }
                
               
            }
        }

        public void WriteLog(object o) {
            WriteLog(o.ToString());
        }

        public void ClearLog() {
            lock (this) {
                File.Create(pathLog);
            }
        }

        ~SystemLoger() {
            WriteLog("Сессия закрыта");
        }
    }
}

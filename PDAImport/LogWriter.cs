using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace PDAImport
{

    public class LogWriter
    {
        public void LogWrite(string filePath, string fileName, string logMessage)
        {
            string fullFilePath = string.Empty;
            FileStream fs = null;
            DirectoryInfo logDirInfo = null;
            FileInfo logFileInfo;
            string fullFileName = string.Empty;
            fullFilePath = filePath + System.DateTime.Today.ToString("yyyyMMdd");
            fullFileName = fullFilePath + "\\" + fileName;

            logFileInfo = new FileInfo(fullFileName);
            logDirInfo = new DirectoryInfo(logFileInfo.DirectoryName);
            if (!logDirInfo.Exists)
                logDirInfo.Create();

            if (!File.Exists(fullFileName)) //No File? Create
            {
                fs = File.Create(fullFileName);
                fs.Close();
            }

            try
            {
                using (StreamWriter w = File.AppendText(fullFileName))
                    AppendLog(logMessage, w);
            }
            catch (Exception ex)
            {
                if (Program.sMode != "silent")
                {
                    Console.WriteLine(ex.Message);
                }
            }

        }

        private static void AppendLog(string logMessage, TextWriter txtWriter)
        {
            try
            {
//                txtWriter.Write("\r\nLog Entry : ");
//                txtWriter.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
//                txtWriter.WriteLine("  :");
                txtWriter.WriteLine("{0}", logMessage);
//                txtWriter.WriteLine("-------------------------------");
            }
            catch (Exception ex)
            {
            }
        }
    }
}

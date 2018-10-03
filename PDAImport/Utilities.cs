using System;
using System.Globalization;
using System.IO;
using System.Linq;

namespace PDAImport
{

    public static class Utilities
    {
        public static int IsInArray(string stringToBeFound, string[][] arr)
        {
            for (int i = 0; i < arr.GetLength(0); i++)
                if (stringToBeFound.Equals(arr[i][0]))
                {
                    return i;
                }
            return -1;
        }

        public static void MoveFile(string fullFileName, string TargetPath, string txtEmail, string txtEmailcc)

        {

            DirectoryInfo logDirInfo = null;
            FileInfo logFileInfo;

            string tFullFileName = TargetPath + System.DateTime.Today.ToString("yyyyMMdd") + "\\" + Path.GetFileName(fullFileName);

            logFileInfo = new FileInfo(fullFileName);
            logDirInfo = new DirectoryInfo(logFileInfo.DirectoryName);
            if (!logDirInfo.Exists)
                logDirInfo.Create();

            try
            {
                if (!File.Exists(fullFileName))
                {
                    // This statement ensures that the file is created,
                    // but the handle is not kept.
                    using (FileStream fs = File.Create(fullFileName)) { }
                }

                // Ensure that the target does not exist.
                if (File.Exists(tFullFileName))
                    File.Delete(tFullFileName);

                // Move the file.
                File.Move(fullFileName, tFullFileName);
                //Console.WriteLine("{0} was moved to {1}.", path, path2);

                // See if the original exists now.
                if (File.Exists(fullFileName))
                {
                    //Console.WriteLine("The original file still exists, which is unexpected.");
                }
                else
                {
                    //Console.WriteLine("The original file no longer exists, which is expected.");
                }

            }
            catch (Exception e)
            {
                string message = "Could not move file - " + fullFileName + " to " + tFullFileName;
                message = message + Environment.NewLine + e.Message;
                email.Sendemail("PDA issue moving file", message, fullFileName, txtEmail, txtEmailcc, null, null);
            }
        }

        public static void CopyFile(string sourcePath, string sourceFile, string targetPath, string txtEmail, string txtEmailcc)

        {
            targetPath = targetPath + System.DateTime.Today.ToString("yyyyMMdd");
            string tFullFileName = targetPath + "\\" + Path.GetFileName(sourceFile);
            string sFullFileName = sourcePath + System.DateTime.Today.ToString("yyyyMMdd") + "\\" + Path.GetFileName(sourceFile);

            try
            {
                if (!System.IO.Directory.Exists(targetPath))
                    System.IO.Directory.CreateDirectory(targetPath);

                // Copy the file and overwrite
                System.IO.File.Copy(sFullFileName, tFullFileName,true);

                // See if the original exists now.
                if (File.Exists(tFullFileName))
                {
                    //Console.WriteLine("The original file still exists, which is unexpected.");
                }
                else
                {
                    //Console.WriteLine("The original file no longer exists, which is expected.");
                }
            }
            catch (Exception e)
            {
                string message = "Could not copy file - " + sFullFileName + " to " + tFullFileName;
                message = message + Environment.NewLine + e.Message;
                email.Sendemail("PDA issue moving file", message, tFullFileName, txtEmail, txtEmailcc, null, null);
            }
        }

        public static int CheckCalendar(string calendarPath, string sLoc, int iLoc)
        {
            // iLoc = 0: Do not process any files           00000000
            // iLoc = 1: Toronto                            00000001
            // iLoc = 2: Montreal                           00000010
            // iLoc = 3: Toronto and Montreal               00000011
            // iLoc = 4: Vancouver                          00000100
            // iLoc = 5: Toronto and Vancouver              00000101
            // iLoc = 6: Vancouver and Montreal             00000110
            // iLoc = 7: Vancouver, Toronto and Montreal    00000111
            // iLoc = 8: Calgary                            00001000
            // iLoc = 9: Calgary and Toronto                00001001
            // iLoc = 10: Calgary, Montreal and Toronto     00001011
            // iLoc = 11: Calgary, Vancouver and Montreal   00001100
            // iLoc = 12: Calgary, Vancouver, Montreal and Toronto 00001101
            // 1 = Toronto
            // 2 = Montreal
            // 4 = Vancouver
            // 8 = Calgary

            string emailBadData = string.Empty;
            string emailBadDatacc = string.Empty;

            TimeSpan start = new TimeSpan(0, 0, 0);
            TimeSpan end = new TimeSpan(0, 0, 0);
            TimeSpan now = new TimeSpan(0, 0, 0);

            try
            {
                string newsLoc = string.Empty;
                string[] lines = File.ReadAllLines(calendarPath);
                string[][] arrayDates = lines.Select(line => line.Split(',').ToArray()).ToArray();
                string sMthDay = DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd");

                foreach (string[] dateLine in arrayDates)
                {
                    if (dateLine[0] == sMthDay)
                    {
                        if (dateLine[1].ToUpper() == "YES")
                        {
                            if (Program.sLoc == "TOR")
                            {
                                start = new TimeSpan(Convert.ToInt32(dateLine[2].Split(':')[0]), Convert.ToInt32(dateLine[2].Split(':')[1]), 0);
                                end = new TimeSpan(Convert.ToInt32(dateLine[3].Split(':')[0]), Convert.ToInt32(dateLine[3].Split(':')[1]), 0);
                                now = DateTime.Now.TimeOfDay;
                                if ((now >= start) && (now <= end))
                                {
                                    iLoc = iLoc ^ 1;
                                }
                            }

                            if (Program.sLoc == "MTL")
                            {
                                start = new TimeSpan(Convert.ToInt32(dateLine[4].Split(':')[0]), Convert.ToInt32(dateLine[4].Split(':')[1]), 0);
                                end = new TimeSpan(Convert.ToInt32(dateLine[5].Split(':')[0]), Convert.ToInt32(dateLine[5].Split(':')[1]), 0);
                                now = DateTime.Now.TimeOfDay;
                                if ((now >= start) && (now <= end))
                                {
                                    iLoc = iLoc ^ 2;
                                }
                            }

                            if (Program.sLoc == "VAN")
                            {
                                start = new TimeSpan(Convert.ToInt32(dateLine[2].Split(':')[0]), Convert.ToInt32(dateLine[2].Split(':')[1]), 0);
                                end = new TimeSpan(Convert.ToInt32(dateLine[3].Split(':')[0]), Convert.ToInt32(dateLine[3].Split(':')[1]), 0);
                                now = DateTime.Now.TimeOfDay;
                                if ((now >= start) && (now <= end))
                                {
                                    iLoc = iLoc ^ 4;
                                }
                            }

                            if (Program.sLoc == "CAL")
                            {
                                start = new TimeSpan(Convert.ToInt32(dateLine[4].Split(':')[0]), Convert.ToInt32(dateLine[4].Split(':')[1]), 0);
                                end = new TimeSpan(Convert.ToInt32(dateLine[5].Split(':')[0]), Convert.ToInt32(dateLine[5].Split(':')[1]), 0);
                                now = DateTime.Now.TimeOfDay;
                                if ((now >= start) && (now <= end))
                                {
                                    iLoc = iLoc ^ 8;
                                }
                            }
                        }
                    }
                }

                return (iLoc);
            }
            catch (Exception e)
            {
                string message = "Calendar - " + calendarPath;
                message = message + "missing!";
                message = message + Environment.NewLine + e.Message;

                if ((Utilities.IsBitSet(Program.iLoc, 4)) || (Utilities.IsBitSet(Program.iLoc, 8)))    // Vancouver or Calgary
                {
                    emailBadData = System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data"];
                    emailBadDatacc = System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data_cc"];
                }
                else
                {
                    emailBadData = System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"];
                    emailBadDatacc = System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"];
                }

                email.Sendemail("PDA Missing Calendar", message, calendarPath, emailBadData, emailBadDatacc, null, null);
                return (0);
            }
        }

        /// <summary>
        /// Returns whether the bit at the specified position is set.
        /// </summary>
        /// <typeparam name="T">Any integer type.</typeparam>
        /// <param name="t">The value to check.</param>
        /// <param name="pos">
        /// The position of the bit to check, 0 refers to the least significant bit.
        /// </param>
        /// <returns>true if the specified bit is on, otherwise false.</returns>
        public static bool IsBitSet<T>(this T t, int pos) where T : struct, IConvertible
        {
            var value = t.ToInt64(CultureInfo.CurrentCulture);
            return (value & (1 << pos)) != 0;
        }
    }
}
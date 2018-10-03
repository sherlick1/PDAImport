using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using ACCPAC.Advantage;


namespace PDAImport
{
    class Program
    {
        // The main entry point for the application.

        public static string sourceSystem;
        public static string sPrd;
        public static string sMode;
        public static string sLoc;
        public static int iLoc;
        public static string output;
        public static bool emailSalesRep;
//        public static string userName;
        public static string sTest;
        public static string CalendarPath;

        public static string Move;
        public static bool CheckOverDue;
        public static string DaysOverDue;
        public static string emailOnHold;
        public static string emailOnHoldcc;
        public static string emailBadData;
        public static string emailBadDatacc;
        public static string emailNoQty;
        public static string emailNoQtycc;
        public static string emailShortage;
        public static string emailShortagecc;
        public static string emailAROverdue;
        public static string emailAROverduecc;
        public static string emailAROverduemtl;
        public static string txtOutputFile;
        public static string txtBadDataFile;
        public static string txtNoQtyFile;
        public static string txtOnHOldFile;
        public static string sourcePath;
//        public static string backupPath;
        public static string torbackupPath;
        public static string vanbackupPath;
        public static string errorPath;

        [STAThread]
        static void Main(string[] clArgs)
        {

            // Get the values of the command line in an array
            // Index  Description
            // 0      Full path of executing program with program name
            // 1      First switch in command: -s
            // 2      First value in command: JFC or SSH
            // 3      Second switch in command: -p
            // 4      Second value in command: DEV or PRD
            // 5      Third switch in command: -m
            // 6      Third value in command: silent or interactive
            // 7      Forth switch in command: -l
            // 8      Forth value in command: location being run for

            clArgs = Environment.GetCommandLineArgs();

            //userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;

            Program.sourceSystem = string.Empty;
            Program.sPrd = string.Empty;
            Program.sMode = string.Empty;
            Program.sTest = string.Empty;
            Program.sLoc = string.Empty;
            Program.iLoc = 0;
            Program.Move = string.Empty;
            Program.CheckOverDue = false;
            Program.DaysOverDue = string.Empty;
            Program.emailOnHold = string.Empty;
            Program.emailOnHoldcc = string.Empty;
            Program.emailBadData = string.Empty;
            Program.emailBadDatacc = string.Empty;
            Program.emailNoQty = string.Empty;
            Program.emailNoQtycc = string.Empty;
            Program.emailShortage = string.Empty;
            Program.emailShortagecc = string.Empty;
            Program.emailAROverdue = string.Empty;
            Program.emailAROverduecc = string.Empty;
            Program.emailAROverduemtl = string.Empty;
            Program.txtOutputFile = string.Empty;
            Program.txtBadDataFile = string.Empty;
            Program.txtNoQtyFile = string.Empty;
            Program.txtOnHOldFile = string.Empty;
            Program.sourcePath = string.Empty;
//            Program.backupPath = string.Empty;
            Program.torbackupPath = string.Empty;
            Program.vanbackupPath = string.Empty;
            Program.errorPath = string.Empty;
            Program.CalendarPath = string.Empty;
            Program.output = "printer";
            Program.emailSalesRep = false;

            Program.CalendarPath = System.Configuration.ConfigurationManager.AppSettings["CalendarPath"];

            if (clArgs.Count() == 9)
            {
                for (int i = 1; i <= 7; i += 2)
                {
                    if (clArgs[i] == "-s")
                        Program.sourceSystem = clArgs[i + 1];
                    else if (clArgs[i] == "-p")
                        Program.sPrd = clArgs[i + 1];
                    else if (clArgs[i] == "-m")
                        Program.sMode = clArgs[i + 1];
                    else if (clArgs[i] == "-l")
                        Program.sLoc = clArgs[i + 1];
                }
            }

            Program.Move = System.Configuration.ConfigurationManager.AppSettings["Move"];
            Program.output = System.Configuration.ConfigurationManager.AppSettings["output"];
            Program.CheckOverDue = Convert.ToBoolean(System.Configuration.ConfigurationManager.AppSettings["CheckOverDue"]);
            Program.DaysOverDue = System.Configuration.ConfigurationManager.AppSettings["DaysOverDue"];
            Program.txtOutputFile = "DailyOrderList.csv";
            Program.txtBadDataFile = "PDA_Bad_Data.csv";
            Program.txtNoQtyFile = "PDA_No_Qty.csv";
            Program.txtOnHOldFile = "PDA_On_Hold.csv";
            Program.torbackupPath = System.Configuration.ConfigurationManager.AppSettings["tor_backup_path"];
            Program.vanbackupPath = System.Configuration.ConfigurationManager.AppSettings["van_backup_path"];

            LogWriter log = new LogWriter();

            if (Program.sMode != "Silent")
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());

            }
            else
            {
                Program.iLoc = 0;

                if (Program.sLoc == "TOR")
                {
                    Program.iLoc = Utilities.CheckCalendar(Program.CalendarPath + "TOR_Calendar.txt", "TOR", Program.iLoc);
                    //log.LogWrite(Program.backupPath, Program.txtOutputFile, "This is a test");
                }
                if (Program.sLoc == "TORMTL")
                {
                    Program.iLoc = Utilities.CheckCalendar(Program.CalendarPath + "TOR_Calendar.txt", "TOR", Program.iLoc);
                    Program.iLoc = Utilities.CheckCalendar(Program.CalendarPath + "TOR_Calendar.txt", "MTL", Program.iLoc);
                }

                if (Program.sLoc == "MTL")
                {
                    Program.iLoc = Utilities.CheckCalendar(Program.CalendarPath + "TOR_Calendar.txt", "MTL", Program.iLoc);
                }

                if (Program.sLoc == "VAN")
                {
                    Program.iLoc = Utilities.CheckCalendar(Program.CalendarPath + "VAN_Calendar.txt", "VAN", Program.iLoc);
                }
                
                if (Program.sLoc == "VANCAL")
                {
                    Program.iLoc = Utilities.CheckCalendar(Program.CalendarPath + "VAN_Calendar.txt", "VAN", Program.iLoc);
                    Program.iLoc = Utilities.CheckCalendar(Program.CalendarPath + "VAN_Calendar.txt", "CAL", Program.iLoc);
                }

                if (Program.sLoc == "CAL")
                {
                    Program.iLoc = Utilities.CheckCalendar(Program.CalendarPath + "VAN_Calendar.txt", "CAL", Program.iLoc);
                }

                if (Program.sLoc == "ALL")
                {
                    Program.iLoc = Utilities.CheckCalendar(Program.CalendarPath + "TOR_Calendar.txt", "TOR", Program.iLoc);
                    Program.iLoc = Utilities.CheckCalendar(Program.CalendarPath + "TOR_Calendar.txt", "MTL", Program.iLoc);
                    Program.iLoc = Utilities.CheckCalendar(Program.CalendarPath + "VAN_Calendar.txt", "VAN", Program.iLoc);
                    Program.iLoc = Utilities.CheckCalendar(Program.CalendarPath + "VAN_Calendar.txt", "CAL", Program.iLoc);
                }

                CreateOrder createorder = new CreateOrder();

                if (createorder.PerformPDAImport())
                {
                    if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                        log.LogWrite(Program.torbackupPath, Program.txtOutputFile, "Import Finished Successfully.");

                    if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                        log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, "Import Finished Successfully.");
                }
                else
                {
                    if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                        log.LogWrite(Program.torbackupPath, Program.txtOutputFile, "Import had errors!");

                    if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                        log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, "Import Finished Successfully.");
                }

                if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                    Utilities.CopyFile(Program.torbackupPath, Program.txtOutputFile, System.Configuration.ConfigurationManager.AppSettings["tor_error_path"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"]);

                if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                    Utilities.CopyFile(Program.vanbackupPath, Program.txtOutputFile, System.Configuration.ConfigurationManager.AppSettings["van_error_path"], System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data_cc"]);
            }
        }
    }
}

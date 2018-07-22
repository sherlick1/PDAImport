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
        public static string userName;
        public static string sTest;

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
        public static string backupPath;
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
        Program.backupPath = string.Empty;
        Program.errorPath = string.Empty;

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

            if (sLoc == "VAN")
            {
                Program.Move = System.Configuration.ConfigurationManager.AppSettings["Move"];
                Program.CheckOverDue = Convert.ToBoolean(System.Configuration.ConfigurationManager.AppSettings["CheckOverDue"]);
                Program.DaysOverDue = System.Configuration.ConfigurationManager.AppSettings["DaysOverDue"];
                Program.sourcePath = System.Configuration.ConfigurationManager.AppSettings["van_source_path"];
                Program.backupPath = System.Configuration.ConfigurationManager.AppSettings["van_backup_path"];
                Program.errorPath = System.Configuration.ConfigurationManager.AppSettings["van_error_path"];
                Program.emailOnHold = System.Configuration.ConfigurationManager.AppSettings["van_email_on_hold"];
                Program.emailOnHoldcc = System.Configuration.ConfigurationManager.AppSettings["van_email_on_hold_cc"];
                Program.emailBadData = System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data"];
                Program.emailBadDatacc = System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data_cc"];
                Program.emailNoQty = System.Configuration.ConfigurationManager.AppSettings["van_email_no_qty"];
                Program.emailNoQtycc = System.Configuration.ConfigurationManager.AppSettings["van_email_no_qty_cc"];
                Program.emailShortage = System.Configuration.ConfigurationManager.AppSettings["van_email_shortage"];
                Program.emailShortagecc = System.Configuration.ConfigurationManager.AppSettings["van_email_shortage_cc"];
                Program.emailAROverdue = System.Configuration.ConfigurationManager.AppSettings["van_email_ar_overdue"];
                Program.emailAROverduecc = System.Configuration.ConfigurationManager.AppSettings["van_email_ar_overdue_cc"];
                Program.emailAROverduemtl = System.Configuration.ConfigurationManager.AppSettings["cal_email_ar_overdue"];
            }
            else
            {
                Program.Move = System.Configuration.ConfigurationManager.AppSettings["Move"];
                Program.DaysOverDue = System.Configuration.ConfigurationManager.AppSettings["DaysOverDue"];
                Program.sourcePath = System.Configuration.ConfigurationManager.AppSettings["tor_source_path"];
                Program.backupPath = System.Configuration.ConfigurationManager.AppSettings["tor_backup_path"];
                Program.errorPath = System.Configuration.ConfigurationManager.AppSettings["tor_error_path"];
                Program.emailOnHold = System.Configuration.ConfigurationManager.AppSettings["tor_email_on_hold"];
                Program.emailOnHoldcc = System.Configuration.ConfigurationManager.AppSettings["tor_email_on_hold_cc"];
                Program.emailBadData = System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"];
                Program.emailBadDatacc = System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"];
                Program.emailNoQty = System.Configuration.ConfigurationManager.AppSettings["tor_email_no_qty"];
                Program.emailNoQtycc = System.Configuration.ConfigurationManager.AppSettings["tor_email_no_qty_cc"];
                Program.emailShortage = System.Configuration.ConfigurationManager.AppSettings["tor_email_shortage"];
                Program.emailShortagecc = System.Configuration.ConfigurationManager.AppSettings["tor_email_shortage_cc"];
                Program.emailAROverdue = System.Configuration.ConfigurationManager.AppSettings["tor_email_ar_overdue"];
                Program.emailAROverduecc = System.Configuration.ConfigurationManager.AppSettings["tor_email_ar_overdue_cc"];
                Program.emailAROverduemtl = System.Configuration.ConfigurationManager.AppSettings["mtl_email_ar_overdue"];
            }

            Program.txtOutputFile = "DailyOrderList.csv";
            Program.txtBadDataFile = "PDA_Bad_Data.csv";
            Program.txtNoQtyFile = "PDA_No_Qty.csv";
            Program.txtOnHOldFile = "PDA_On_Hold.csv";

            LogWriter log = new LogWriter();
            log.LogWrite(Program.backupPath, Program.txtOutputFile, "This is a test");
            //log.LogWrite(@"c:\temp\", "test.txt", "This is a test");


            if (Program.sMode != "silent")
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
            }
            else
            {
                CreateOrder createorder = new CreateOrder();

                if (createorder.PerformPDAImport(sLoc))
                {
                    log.LogWrite(Program.backupPath, Program.txtOutputFile, "Import Finished Successfully.");
                }
                else
                {
                    log.LogWrite(Program.backupPath, Program.txtOutputFile, "Import had errors!");
                }
            }
        }
    }
}

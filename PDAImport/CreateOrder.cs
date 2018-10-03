using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using ACCPAC.Advantage;
using System.Text.RegularExpressions;

namespace PDAImport
{
    public partial class CreateOrder
    {

        /**********************************************
         * Declarations
         **********************************************/
        public event MyHandler1 Event1;

        // Created by Stephen Herlick - August, 2018
        // Based on this article
        // https://smist08.wordpress.com/2010/11/07/entering-orders-and-headerdetail-views/
        // 
        // 
        // Macro to enter orders into O/E
        // 
        // Global variables for database links and views that are needed in most routines.
        // 
        private ACCPAC.Advantage.Session sageSession;
        private ACCPAC.Advantage.DBLink mDBLinkCmpRW;

        private ACCPAC.Advantage.View OEORD1header;
        private ACCPAC.Advantage.ViewFields OEORD1headerFields;
        private ACCPAC.Advantage.View OEORD1detail1;
        private ACCPAC.Advantage.ViewFields OEORD1detail1Fields;
        private ACCPAC.Advantage.View OEORD1detail2;
        private ACCPAC.Advantage.ViewFields OEORD1detail2Fields;
        private ACCPAC.Advantage.View OEORD1detail3;
        private ACCPAC.Advantage.ViewFields OEORD1detail3Fields;
        private ACCPAC.Advantage.View OEORD1detail4;
        private ACCPAC.Advantage.ViewFields OEORD1detail4Fields;
        private ACCPAC.Advantage.View OEORD1detail5;
        private ACCPAC.Advantage.ViewFields OEORD1detail5Fields;
        private ACCPAC.Advantage.View OEORD1detail6;
        private ACCPAC.Advantage.ViewFields OEORD1detail6Fields;
        private ACCPAC.Advantage.View OEORD1detail7;
        private ACCPAC.Advantage.ViewFields OEORD1detail7Fields;
        private ACCPAC.Advantage.View OEORD1detail8;
        private ACCPAC.Advantage.ViewFields OEORD1detail8Fields;
        private ACCPAC.Advantage.View OEORD1detail9;
        private ACCPAC.Advantage.ViewFields OEORD1detail9Fields;
        private ACCPAC.Advantage.View OEORD1detail11;
        private ACCPAC.Advantage.ViewFields OEORD1detail11Fields;
        private ACCPAC.Advantage.View OEORD1detail10;
        private ACCPAC.Advantage.ViewFields OEORD1detail10Fields;
        private ACCPAC.Advantage.View OEORD1detail12;
        private ACCPAC.Advantage.ViewFields OEORD1detail12Fields;

//        private ACCPAC.Advantage.View OESHI1header;
//        private ACCPAC.Advantage.View OESHI1detail1;
//        private ACCPAC.Advantage.View OESHI1detail2;
//        private ACCPAC.Advantage.View OESHI1detail3;
//        private ACCPAC.Advantage.View OESHI1detail4;
//        private ACCPAC.Advantage.View OESHI1detail5;
//        private ACCPAC.Advantage.View OESHI1detail6;
//        private ACCPAC.Advantage.View OESHI1detail7;
//        private ACCPAC.Advantage.View OESHI1detail8;
//        private ACCPAC.Advantage.View OESHI1detail9;
//        private ACCPAC.Advantage.View OESHI1detail10;
//        private ACCPAC.Advantage.View OESHI1detail11;
//        private ACCPAC.Advantage.View OESHI1detail12;

        private ACCPAC.Advantage.View ICITEM_T1header;
        private ACCPAC.Advantage.ViewFields ICITEM_T1headerFields;
        private ACCPAC.Advantage.View ICITEM_T1detail1;
        private ACCPAC.Advantage.ViewFields ICITEM_T1detail1Fields;
        private ACCPAC.Advantage.View ICITEM_T1detail2;
        private ACCPAC.Advantage.ViewFields ICITEM_T1detail2Fields;
        private ACCPAC.Advantage.View ICITEM_T1detail3;
        private ACCPAC.Advantage.ViewFields ICITEM_T1detail3Fields;
        private ACCPAC.Advantage.View ICITEM_T1detail4;
        private ACCPAC.Advantage.ViewFields ICITEM_T1detail4Fields;
        private ACCPAC.Advantage.View ICITEM_T1detail5;
        //private ACCPAC.Advantage.ViewFields ICITEM_T1detail5Fields;
        private ACCPAC.Advantage.View ICITEM_T1detail6;
        private ACCPAC.Advantage.ViewFields ICITEM_T1detail6Fields;
        private ACCPAC.Advantage.View ICITEM_T1detail7;
        //private ACCPAC.Advantage.ViewFields ICITEM_T1detail7Fields;
        private ACCPAC.Advantage.View ICITEM_T1detail8;
        private ACCPAC.Advantage.ViewFields ICITEM_T1detail8Fields;
        private ACCPAC.Advantage.View ICITEM_T1detail9;
        private ACCPAC.Advantage.ViewFields ICITEM_T1detail9Fields;
        private ACCPAC.Advantage.View ICITEM_T1detail10;
        private ACCPAC.Advantage.ViewFields ICITEM_T1detail10Fields;
        private ACCPAC.Advantage.View ICITEM_T1detail11;
        private ACCPAC.Advantage.ViewFields ICITEM_T1detail11Fields;

        private ACCPAC.Advantage.View ICILOC1;
        private ACCPAC.Advantage.ViewFields ICILOC1Fields;

        public string[][] arrayofSalesReps;

        // User-defined constants
        public const string numFormat = "#,##0.00";
        public const string qtyFormat = "#,##0.0000";
        public const string doubleQuote = "\"";
        public const string singleQuote = "\'";

        public struct Record
        {
            public readonly string[] Row;
            public string this[int index] => Row[index];
            public Record(string[] row)
            {
                Row = row;
            }
        }

        public string OpenAndComposeOrderViews()
        {

            OEORD1header = mDBLinkCmpRW.OpenView("OE0520");
            OEORD1detail1 = mDBLinkCmpRW.OpenView("OE0500");
            OEORD1detail2 = mDBLinkCmpRW.OpenView("OE0740");
            OEORD1detail3 = mDBLinkCmpRW.OpenView("OE0180");
            OEORD1detail4 = mDBLinkCmpRW.OpenView("OE0526");
            OEORD1detail5 = mDBLinkCmpRW.OpenView("OE0522");
            OEORD1detail6 = mDBLinkCmpRW.OpenView("OE0508");
            OEORD1detail7 = mDBLinkCmpRW.OpenView("OE0507");
            OEORD1detail8 = mDBLinkCmpRW.OpenView("OE0501");
            OEORD1detail9 = mDBLinkCmpRW.OpenView("OE0502");
            OEORD1detail10 = mDBLinkCmpRW.OpenView("OE0504");
            OEORD1detail11 = mDBLinkCmpRW.OpenView("OE0506");
            OEORD1detail12 = mDBLinkCmpRW.OpenView("OE0503");

            OEORD1header.Compose(new ACCPAC.Advantage.View[] { OEORD1detail1, null, OEORD1detail3, OEORD1detail2, OEORD1detail4, OEORD1detail5 });
            OEORD1headerFields = OEORD1header.Fields;
            OEORD1detail1.Compose(new ACCPAC.Advantage.View[] { OEORD1header, OEORD1detail8, OEORD1detail12, OEORD1detail9, OEORD1detail6, OEORD1detail7 });
            OEORD1detail1Fields = OEORD1detail1.Fields;
            OEORD1detail2.Compose(new ACCPAC.Advantage.View[] { OEORD1header });
            OEORD1detail2Fields = OEORD1detail2.Fields;
            OEORD1detail3.Compose(new ACCPAC.Advantage.View[] { OEORD1header, OEORD1detail1 });
            OEORD1detail3Fields = OEORD1detail3.Fields;
            OEORD1detail4.Compose(new ACCPAC.Advantage.View[] { OEORD1header });
            OEORD1detail4Fields = OEORD1detail4.Fields;
            OEORD1detail5.Compose(new ACCPAC.Advantage.View[] { OEORD1header });
            OEORD1detail5Fields = OEORD1detail5.Fields;
            OEORD1detail6.Compose(new ACCPAC.Advantage.View[] { OEORD1detail1 });
            OEORD1detail6Fields = OEORD1detail6.Fields;
            OEORD1detail7.Compose(new ACCPAC.Advantage.View[] { OEORD1detail1 });
            OEORD1detail7Fields = OEORD1detail7.Fields;
            OEORD1detail8.Compose(new ACCPAC.Advantage.View[] { OEORD1detail1 });
            OEORD1detail8Fields = OEORD1detail8.Fields;
            OEORD1detail9.Compose(new ACCPAC.Advantage.View[] { OEORD1detail1, OEORD1detail10, OEORD1detail11 });
            OEORD1detail9Fields = OEORD1detail9.Fields;
            OEORD1detail10.Compose(new ACCPAC.Advantage.View[] { OEORD1detail9 });
            OEORD1detail10Fields = OEORD1detail10.Fields;
            OEORD1detail11.Compose(new ACCPAC.Advantage.View[] { OEORD1detail9 });
            OEORD1detail11Fields = OEORD1detail11.Fields;
            OEORD1detail12.Compose(new ACCPAC.Advantage.View[] { OEORD1detail1 });
            OEORD1detail12Fields = OEORD1detail12.Fields;

            OEORD1headerFields.FieldByName("DRIVENBYUI").SetValue("0", false);

            return "OK";

        }

        public void CleanUpOrderViews()
        {

            OEORD1header.Dispose();
            OEORD1detail1.Dispose();
            OEORD1detail2.Dispose();
            OEORD1detail3.Dispose();
            OEORD1detail4.Dispose();
            OEORD1detail5.Dispose();
            OEORD1detail6.Dispose();
            OEORD1detail7.Dispose();
            OEORD1detail8.Dispose();
            OEORD1detail9.Dispose();
            OEORD1detail10.Dispose();
            OEORD1detail11.Dispose();
            OEORD1detail12.Dispose();
            OEORD1headerFields.Dispose();
            OEORD1detail1Fields.Dispose();
            OEORD1detail2Fields.Dispose();
            OEORD1detail3Fields.Dispose();
            OEORD1detail4Fields.Dispose();
            OEORD1detail5Fields.Dispose();
            OEORD1detail6Fields.Dispose();
            OEORD1detail7Fields.Dispose();
            OEORD1detail8Fields.Dispose();
            OEORD1detail9Fields.Dispose();
            OEORD1detail10Fields.Dispose();
            OEORD1detail11Fields.Dispose();
            OEORD1detail12Fields.Dispose();

        }

        public Boolean PerformPDAImport()
        {

            LogWriter log = new LogWriter();

            List<string> ListofFiles = new List<string>();
            ListofFiles = GetFileNames();

            if (ListofFiles.Count > 0)
            {

                if (Program.sMode != "Silent")
                {
                    FireEvent("Logging into Sage");
                }

                if (SageLogin())
                {
                    if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                        //log.LogWrite(Program.torbackupPath, Program.txtOutputFile, "Successfully logged into Sage.");

                    if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                        //log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, "Successfully logged into Sage.");


                    if (Program.sMode != "Silent")
                    {
                        FireEvent("Successfully logged in.");
                    }

                    arrayofSalesReps = GetSalesReps();

                    if (Program.sMode != "Silent")
                    {
                        FireEvent("Processing files....");
                    }


                    foreach (string fullFileName in ListofFiles)
                    {

                        if (Program.sMode != "Silent")
                        {
                            FireEvent("Processing file " + fullFileName);
                        }

                        /* https://stackoverflow.com/questions/2081418/parsing-csv-files-in-c-with-header
                           Using the generic ReadCSVFile code at the end */

                        string status = ProcessFile(fullFileName, Program.sLoc);
                    }
                    return true;
                }
                else
                {
                    if (Program.sMode != "Silent")
                    {
                        FireEvent("log into Sage failed.");
                    }
                    if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                        log.LogWrite(Program.torbackupPath, Program.txtOutputFile, "Failed to log into Sage.");

                    if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                        log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, "Failed to log into Sage.");

                    return false;
                }
            }

            if (Program.sMode != "Silent")
            {
                FireEvent("No files to process.");
            }

            if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                log.LogWrite(Program.torbackupPath, Program.txtOutputFile, "No files to process.");

            if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Toronto or Montreal
                log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, "No files to process.");

            return true;
        }

        // Login into Sage
        public Boolean SageLogin()
        {
            try
            {
                short periods;
                bool bActive, bRet;
                string year;

                // Create, initialize and open a session.
                sageSession = new Session();
                sageSession.Init("", "XX", "XX1000", "63A");
                sageSession.Open("ADMIN", "RTFM", "JFJDAT", DateTime.Today, 0);

                // Open a database link.
                mDBLinkCmpRW = sageSession.OpenDBLink(DBLinkType.Company, DBLinkFlags.ReadWrite);

                ACCPAC.Advantage.FiscalCalendar fiscCal = mDBLinkCmpRW.FiscalCalendar;

                bRet = fiscCal.GetPeriod(new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day), out periods, out year, out bActive);
                return true;
            }
            catch (Exception ex)
            {
                string message = "Cannot log into Sage";

                if (Program.sMode != "Silent")
                {
                    FireEvent("No files to process.");
                }
                email.Sendemail("PDA macro could not log into SAGE", message, null, Program.emailBadData, Program.emailBadDatacc, null, null);
                MyErrorHandler(ex, message);
                return false;
            }
        }

        List<string> GetFileNames()
        {
            List<string> ListofFiles = new List<string>();

            // The following could work but the issue is the full path name is returned so the reg ex pattern
            // check needs to account for the full path name which is variable.
            //                var Files = Directory.GetFiles(@"c:\Files\Clients\JFC\PDA Files\TOR\sfa\")
            //                       .Where(Path => reg.IsMatch(Path))
            //                      .ToList();

            if (Program.sourceSystem == "JFC")
            {
                if ((Program.sLoc == "VAN") || (Program.sLoc == "CAL") || (Program.sLoc == "VANCAL") || (Program.sLoc == "ALL"))
                {
                    ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"D:\KMS\PDA\VA\var\spool\jetsso\", "*", SearchOption.AllDirectories));
                    // ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"D:\KMS\PDA\VA\var\spool\jetsso\SFA\", "*", SearchOption.AllDirectories));
                    // ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"D:\KMS\PDA\VA\var\spool\jetsso\PDA\", "*", SearchOption.AllDirectories));
                    // ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"D:\PDA_Files\VA\REPROCESS\", "*", SearchOption.AllDirectories));
                }

                if ((Program.sLoc == "TOR") || (Program.sLoc == "MTL") || (Program.sLoc == "TORMTL") || (Program.sLoc == "ALL"))
                {
                    ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"D:\KMS\PDA\TR\var\spool\jetsso\", "*", SearchOption.AllDirectories));
                    // ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"D:\KMS\PDA\TR\var\spool\jetsso\SFA\", "*", SearchOption.AllDirectories));
                    // ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"D:\KMS\PDA\TR\var\spool\jetsso\PDA\", "*", SearchOption.AllDirectories));
                    // ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"D:\PDA_Files\TR\REPROCESS\", "*", SearchOption.AllDirectories));
                }
            }
            else
            {
                if ((Program.sLoc == "VAN") || (Program.sLoc == "CAL") || (Program.sLoc == "VANCAL") || (Program.sLoc == "ALL"))
                {
                    ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"c:\Files\Clients\JFC\PDA Files\VAN\", "*", SearchOption.AllDirectories));
                    // ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"c:\Files\Clients\JFC\PDA Files\VAN\SFA\", "*", SearchOption.AllDirectories));
                    // ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"c:\Files\Clients\JFC\PDA Files\VAN\PDA\", "*", SearchOption.AllDirectories));
                    // ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"c:\Files\Clients\JFC\PDA Files\REPROCESS", "*", SearchOption.AllDirectories));
                }

                if ((Program.sLoc == "TOR") || (Program.sLoc == "MTL") || (Program.sLoc == "TORMTL") || (Program.sLoc == "ALL"))
                {
                    ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"c:\Files\Clients\JFC\PDA Files\TOR\", "*", SearchOption.AllDirectories));
                    // ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"c:\Files\Clients\JFC\PDA Files\TOR\SFA\", "*", SearchOption.AllDirectories));
                    // ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"c:\Files\Clients\JFC\PDA Files\TOR\PDA\", "*", SearchOption.AllDirectories));
                    // ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"c:\Files\Clients\JFC\PDA Files\REPROCESS\", "*", SearchOption.AllDirectories));
                }
            }

            return ListofFiles;
        }

        private List<string> FilterFiles(List<string> ListofFiles, string[] ArrayofFiles)
        {
            FileInfo fileInfo;
            Regex reg = new Regex(@"^[frs]\d{8,10}\.\d{4}$");

            foreach (string fullFileName in ArrayofFiles)
            {
                fileInfo = new FileInfo(fullFileName);
                if (reg.IsMatch(fileInfo.Name))
                    ListofFiles.Add(fullFileName);
            }

            return ListofFiles;
        }

        string[][] GetSalesReps()
        {

            string salesRepsPath;

            if (Program.sourceSystem == "JFC")
            {
                if (Program.sLoc == "VAN")
                {
                    salesRepsPath = @"D:\PDA_Files\SETTINGS\Salesreps.ini";
                }
                else
                {
                    salesRepsPath = @"D:\PDA_Files\SETTINGS\Salesreps.ini";
                }
            }
            else
            {
                if (Program.sLoc == "VAN")
                {
                    salesRepsPath = @"c:\Files\Clients\JFC\PDA Process\Salesreps.ini";
                }
                else
                {
                    salesRepsPath = @"c:\Files\Clients\JFC\PDA Process\Salesreps.ini";
                }
            }

            try
            {

                string[] lines = File.ReadAllLines(salesRepsPath);
                string[][] locArrayofSalesReps = lines.Select(line => line.Split(',').ToArray()).ToArray();

                return locArrayofSalesReps;
            }

            catch (Exception e)
            {
                string message = "Sales rep file - " + salesRepsPath + Environment.NewLine;
                message = message + "is missing!" + Environment.NewLine;
                message = message + e.Message;
                email.Sendemail("PDA Missing Sales Rep File", message, salesRepsPath, Program.emailBadData, Program.emailBadDatacc, null, null);
                return (null);
            }
        }

        private string CheckSalesPerson(string sSalesPerson, string custNo, string reference, string fullFileName, string sendeMailStatus)
        {
            Boolean fnd;
            String message;

            ACCPAC.Advantage.View arSalesPsn;

            try
            {
                arSalesPsn = mDBLinkCmpRW.OpenView("AR0018");
                arSalesPsn.Fields.FieldByName("CODESLSP").SetValue(sSalesPerson, false);
                arSalesPsn.Read(false);

                fnd = arSalesPsn.Exists;

                if (fnd)
                {
                    if (Convert.ToInt32(arSalesPsn.Fields.FieldByName("SWACTV").Value) == 0)
                    {
                        message = "PDA Bad Data Import Error";
                        message = message + Environment.NewLine + "File - " + fullFileName;
                        message = message + Environment.NewLine + "SalesPerson - " + sSalesPerson;
                        message = message + Environment.NewLine + "SalesPerson is inactive.";
                        message = message + Environment.NewLine + "Reference - " + reference;
                        message = message + Environment.NewLine + "Customer - " + custNo;
                        email.Sendemail("PDA Bad SalesPerson", message, fullFileName, System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"], sSalesPerson, arrayofSalesReps);

                        arSalesPsn.Dispose();
                        return ("BAD");
                    }
                    else
                        arSalesPsn.Dispose();
                    return "OK";
                }
                else
                {

                    message = "PDA Bad Data Import Error";
                    message = message + Environment.NewLine + "File - " + fullFileName;
                    message = message + Environment.NewLine + "SalesPerson - " + sSalesPerson;
                    message = message + Environment.NewLine + "SalesPerson is missing.";
                    message = message + Environment.NewLine + "Reference - " + reference;
                    message = message + Environment.NewLine + "Customer - " + custNo;
                    email.Sendemail("PDA Bad SalesPerson", message, fullFileName, System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"], sSalesPerson, arrayofSalesReps);
                    arSalesPsn.Dispose();
                    return "BAD";
                }
                }
            catch (Exception ex)
            {
                message = "PDA Bad Data Import Error in CheckSalesPerson";
                message = message + Environment.NewLine + "File - " + fullFileName;
                message = message + Environment.NewLine + "SalesPerson - " + sSalesPerson;
                message = message + Environment.NewLine + ex.Message;
                message = message + Environment.NewLine + "Reference - " + reference;
                message = message + Environment.NewLine + "Customer - " + custNo;
                MyErrorHandler(ex, message);
                return ("BAD");
            }

        }

        private string CheckCustomer(string custNo, string fullFileName, ref string custName, bool printShipment, bool printConfirmation, string sendeMailStatus)
        {
            Boolean fnd;
            string message;
            string status = "OK";

            ACCPAC.Advantage.View arCustomerHeader;
            ACCPAC.Advantage.View arCustomerDetail;

            try
            {
                arCustomerHeader = mDBLinkCmpRW.OpenView("AR0024");
                arCustomerDetail = mDBLinkCmpRW.OpenView("AR0400");
                arCustomerHeader.Fields.FieldByName("IDCUST").SetValue(custNo, false);
                arCustomerHeader.Read(false);

                fnd = arCustomerHeader.Exists;

                if (fnd)
                {
                    custName = Convert.ToString(arCustomerHeader.Fields.FieldByName("NAMECUST").Value);

                    if (Convert.ToBoolean(arCustomerHeader.Fields.FieldByName("SWHOLD").Value))
                    {
                        printShipment = false;
                        status = "ONHOLD";
                    }

                    printConfirmation = false;

                    if (Program.CheckOverDue)
                    {
                        ACCPAC.Advantage.View csQry;
                        csQry = mDBLinkCmpRW.OpenView("CS0120");

                        string sSQL;
                        //Below is the query that can be passed to the CSQRY
                        sSQL = "SELECT [IDCUST], ";
                        sSQL = sSQL + "[IDINVC], ";
                        sSQL = sSQL + "[TRXTYPETXT], ";
                        sSQL = sSQL + "[DATEINVC], ";
                        sSQL = sSQL + "substring(cast([DATEINVC] as CHAR(8)),1,4) As YEAR, ";
                        sSQL = sSQL + "substring(cast([DATEINVC] as CHAR(8)),5,2) As MONTH, ";
                        sSQL = sSQL + "substring(cast([DATEINVC] as CHAR(8)),7,2) As DAY, ";
                        sSQL = sSQL + "cast(substring(cast([DATEINVC] as CHAR(8)),1,4) + substring(cast([DATEINVC] as CHAR(8)),5,2) + substring(cast([DATEINVC] as CHAR(8)),7,2) AS DATE) AS DUEINVC, ";
                        //  sSQL = sSQL + ",DATEDIFF(DAY,GETDATE(),cast(substring(cast([DATEINVC] as CHAR(8)),1,4) + substring(cast([DATEINVC] as CHAR(8)),5,2) + substring(cast([DATEINVC] as CHAR(8)),7,2) AS DATE)) AS DAYS ";
                        sSQL = sSQL + "DATEDIFF(DAY,cast(substring(cast([DATEINVC] as CHAR(8)),1,4) + substring(cast([DATEINVC] as CHAR(8)),5,2) + substring(cast([DATEINVC] as CHAR(8)),7,2) AS DATE),GETDATE()) AS DAYS ";
                        if (Program.sourceSystem == "JFC")
                        {
                            if (Program.sPrd == "PRD")
                                sSQL = sSQL + "From [JFJDAT].[dbo].[AROBL] ";
                            else
                                sSQL = sSQL + "From [REFJFJ].[dbo].[AROBL] ";
                        }
                        else
                            sSQL = sSQL + "From [JFJDAT].[dbo].[AROBL] ";

                        sSQL = sSQL + "WHERE [IDCUST] = " + singleQuote + custNo + singleQuote;
                        sSQL = sSQL + " and [IDINVC] <> ''";
                        sSQL = sSQL + " and ([TRXTYPETXT] IN (1,2))";
                        sSQL = sSQL + " and [SWPAID] = 0;";

                        csQry.Cancel();
                        csQry.Browse(sSQL, true);
                        csQry.InternalSet(256);

                        bool bcontinue = true;

                        while (true == csQry.Fetch(false) && bcontinue)
                        {
                            int resultval = Convert.ToInt32(csQry.Fields[8].Value.ToString());
                            if (resultval > Convert.ToInt32(Program.DaysOverDue))
                            {
                                //PrintShipment = False
                                if (status == "ONHOLD")
                                {
                                    status = "ONHOLD_AND_OVERDUE";
                                    bcontinue = false;
                                }
                                else
                                {
                                    status = "OVERDUE";
                                    bcontinue = false;
                                }
                            }
                        }
                    }
                    else
                        status = "OK";
                }
                else
                    status = "BAD";

                arCustomerHeader.Dispose();
                arCustomerDetail.Dispose();
                return (status);
            }

            catch (Exception ex)
            {
                message = "PDA Bad Data Import Error in CheckCustomer";
                message = message + Environment.NewLine + ex.Message;
                message = message + Environment.NewLine + "File - " + fullFileName;
                message = message + Environment.NewLine + "Customer - " + custNo;
                MyErrorHandler(ex, message);
                return ("BAD");
            }

        }

        private string CheckShipToCustomer(string custNo, string fullFileName, string custName, string shipTo, string sendeMailStatus)
        {
            Boolean fnd;
            string message;
            string status = "OK";

            ACCPAC.Advantage.View arShipToCustomerHeader;
            ACCPAC.Advantage.View arShipToCustomerDetail;

            try
            {
                arShipToCustomerHeader = mDBLinkCmpRW.OpenView("AR0023");
                arShipToCustomerDetail = mDBLinkCmpRW.OpenView("AR0412");
                arShipToCustomerHeader.Fields.FieldByName("IDCUST").SetValue(custNo, false);
                arShipToCustomerHeader.Fields.FieldByName("IDCUSTSHPT").SetValue(shipTo, false);
                arShipToCustomerHeader.Read(false);

                fnd = arShipToCustomerHeader.Exists;

                if (fnd)
                    status = "OK";
                else
                    status = "BAD";

                arShipToCustomerHeader.Dispose();
                arShipToCustomerDetail.Dispose();
                return (status);
            }

            catch (Exception ex)
            {
                message = "PDA Bad Data Import Error in CheckShipToCustomer";
                message = message + Environment.NewLine + ex.Message;
                message = message + Environment.NewLine + "File - " + fullFileName;
                message = message + Environment.NewLine + "Customer - " + custNo;
                MyErrorHandler(ex, message);
                return ("BAD");
            }

        }

        private void CheckItem(string itemNo, string fullFileName, string location, double qtyOrdered, string orderUnit, ref string itemStatus, bool postOrder, ref double qtyShipped, List<ItemsShippedLine> itemsShipped, ref decimal priUntPrc)
        {

//            Boolean temp;
//            int lFnd;
            LogWriter log = new LogWriter();

            if (OpenAndComposeItemViews())
            {
                float qtyOnHand = 0;
                double convQtyOnHand = 0;
                string itemDescription = string.Empty;

                try
                {
                    ICITEM_T1headerFields.FieldByName("ITEMNO").SetValue(itemNo, false);           // Item Number
                    ICITEM_T1header.Read(false);
                    if (ICITEM_T1header.Exists)
                    {
                        bool inActive = Convert.ToBoolean(ICITEM_T1headerFields.FieldByName("INACTIVE").Value);
                        bool stockItem = Convert.ToBoolean(ICITEM_T1headerFields.FieldByName("STOCKITEM").Value);

                        if (!inActive && stockItem)
                        {
                            ICILOC1Fields.FieldByName("ITEMNO").SetValue(itemNo, false);
                            ICILOC1Fields.FieldByName("LOCATION").SetValue(location, false);
                            ICILOC1.Read(false);

                            if (ICILOC1.Exists)
                            {
                                if (Convert.ToBoolean(ICILOC1Fields.FieldByName("ACTIVE").Value))
                                {
                                    ICITEM_T1detail1Fields.FieldByName("ITEMNO").SetValue(itemNo, false);
                                    ICITEM_T1detail1Fields.FieldByName("UNIT").SetValue(orderUnit, false);
                                    ICITEM_T1detail1.Read(false);

                                    if (ICITEM_T1detail1.Exists)
                                    {
                                        qtyOnHand = float.Parse(ICILOC1Fields.FieldByName("AQTYONHAND").Value.ToString(),System.Globalization.CultureInfo.InvariantCulture);

                                        foreach (ItemsShippedLine itemNoFnd in itemsShipped)
                                        {
                                            if (itemNoFnd.itemNo == itemNo)
                                            {
                                                qtyOnHand = qtyOnHand - itemNoFnd.qtyShipped;
                                               // break; // If you only want to find the first instance a break here would be best for your application
                                            }
                                        }

                                        float conversion = float.Parse(ICITEM_T1detail1Fields.FieldByName("CONVERSION").Value.ToString(),System.Globalization.CultureInfo.InvariantCulture);

                                        if (conversion != 1 && qtyOnHand > 0)
                                        {
                                            try
                                            {
                                                convQtyOnHand = (Math.Truncate(qtyOnHand / conversion * 10000)) / 10000;

                                                if (convQtyOnHand < qtyOrdered)
                                                {
                                                    itemStatus = "Not_Enough_Quantity";
                                                    if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                                                        log.LogWrite(Program.torbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + itemNo + " quantity ordered greater than greater on hand in location " + location + ".");

                                                    if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Montreal
                                                        log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + itemNo + " quantity ordered greater than greater on hand in location " + location + ".");

                                                    if (qtyOnHand <= 0)
                                                        qtyShipped = 0;
                                                    else
                                                    {
                                                        qtyShipped = convQtyOnHand;
                                                        itemsShipped.Add(new ItemsShippedLine { itemNo = itemNo, qtyShipped = qtyOnHand, uom = orderUnit });
                                                        priUntPrc = priUntPrc * Convert.ToDecimal(conversion);
                                                    }
                                                }
                                                else
                                                {
                                                    qtyShipped = qtyOrdered;
                                                    convQtyOnHand = (Math.Truncate(qtyOrdered * conversion * 10000)) / 10000;
                                                    itemsShipped.Add(new ItemsShippedLine { itemNo = itemNo, qtyShipped = (float)convQtyOnHand, uom = orderUnit });
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                string message = "CheckItem Error";
                                                message = message + Environment.NewLine + "Item: " + itemNo;
                                                message = message + Environment.NewLine + ex.Message;
                                                MyErrorHandler(ex, message);
                                                CleanUpItemViews();
                                            }

                                        }
                                        else if (qtyOnHand <= 0)
                                        {
                                            qtyOnHand = 0;
                                            itemStatus = "Not_Enough_Quantity";

                                            if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                                                log.LogWrite(Program.torbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + itemNo + " quantity ordered greater than greater on hand in location " + location + ".");

                                            if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                                                log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + itemNo + " quantity ordered greater than greater on hand in location " + location + ".");

                                        }
                                        else
                                        {
                                            if (qtyOnHand >= qtyOrdered)
                                            {
                                                qtyShipped = qtyOrdered;
                                                itemsShipped.Add(new ItemsShippedLine { itemNo = itemNo, qtyShipped = (float)qtyShipped, uom = orderUnit });
                                            }
                                            else
                                            {
                                                qtyShipped = qtyOnHand;
                                                itemsShipped.Add(new ItemsShippedLine { itemNo = itemNo, qtyShipped = (float)qtyShipped, uom = orderUnit });
                                            }
                                        }
                                    }
                                    else
                                    {
                                        itemStatus = "Invalid_Order_Unit";
                                        postOrder = false;

                                        if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                                            log.LogWrite(Program.torbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + itemNo + " order unit " + orderUnit + " is not valid for location " + location + ".");

                                        if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                                            log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + itemNo + " order unit " + orderUnit + " is not valid for location " + location + ".");

                                    }
                                }
                                else
                                {
                                    itemStatus = "Item_Inactive_at_Location";
                                    postOrder = false;

                                    if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                                        log.LogWrite(Program.torbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + itemNo + " item inactive in location " + location + ".");

                                    if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                                        log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + itemNo + " item inactive in location " + location + ".");

                                }
                            }
                            else
                            {
                                itemStatus = "Item_Does_Not_Exist_at_Location";
                                postOrder = false;

                                if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                                    log.LogWrite(Program.torbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + itemNo + " item does not exist in location " + location + ".");

                                if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                                    log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + itemNo + " item does not exist in location " + location + ".");

                            }
                        }
                        else
                        {
                            itemStatus = "Item_inactive_at_location";
                            postOrder = false;

                            if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                                log.LogWrite(Program.torbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + itemNo + " item is not active in location " + location + ".");

                            if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                                log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + itemNo + " item is not active in location " + location + ".");

                        }

                        CleanUpItemViews();
                    }
                    else
                    {
                        itemStatus = "Item_Does_Not_Exist";
                        postOrder = false;

                        if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                            log.LogWrite(Program.torbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + itemNo + " item does not exist.");

                        if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                            log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + itemNo + " item does not exist.");

                    }
                }

                catch (Exception ex)
                {
                    string message = "CheckItem Error";
                    message = message + Environment.NewLine + "Items: " + itemNo;
                    message = message + Environment.NewLine + ex.Message;
                    MyErrorHandler(ex, message);
                    CleanUpItemViews();
                }
            }
        }

        public bool OpenAndComposeItemViews()
        {

            try
            {
                ICITEM_T1header = mDBLinkCmpRW.OpenView("IC0310");
                ICITEM_T1detail1 = mDBLinkCmpRW.OpenView("IC0750");
                ICITEM_T1detail2 = mDBLinkCmpRW.OpenView("IC0330");
                ICITEM_T1detail3 = mDBLinkCmpRW.OpenView("IC0340");
                ICITEM_T1detail4 = mDBLinkCmpRW.OpenView("IC0320");
                ICITEM_T1detail5 = mDBLinkCmpRW.OpenView("IC0210");
                ICITEM_T1detail6 = mDBLinkCmpRW.OpenView("IC0100");
                ICITEM_T1detail7 = mDBLinkCmpRW.OpenView("IC0390");
                ICITEM_T1detail8 = mDBLinkCmpRW.OpenView("IC0313");
                ICITEM_T1detail9 = mDBLinkCmpRW.OpenView("IC0319");
                ICITEM_T1detail10 = mDBLinkCmpRW.OpenView("IC0314");
                ICITEM_T1detail11 = mDBLinkCmpRW.OpenView("IC0312");

                ICITEM_T1header.Compose(new ACCPAC.Advantage.View[] { ICITEM_T1detail1, ICITEM_T1detail2, ICITEM_T1detail3, ICITEM_T1detail4, ICITEM_T1detail5, ICITEM_T1detail6, ICITEM_T1detail7, ICITEM_T1detail8, ICITEM_T1detail9, ICITEM_T1detail10, ICITEM_T1detail11 });
                ICITEM_T1headerFields = ICITEM_T1header.Fields;
                ICITEM_T1detail1.Compose(new ACCPAC.Advantage.View[] { null });
                ICITEM_T1detail1Fields = ICITEM_T1detail1.Fields;
                ICITEM_T1detail2.Compose(new ACCPAC.Advantage.View[] { ICITEM_T1header, null, null });
                ICITEM_T1detail2Fields = ICITEM_T1detail2.Fields;
                ICITEM_T1detail3.Compose(new ACCPAC.Advantage.View[] { ICITEM_T1header, null, null });
                ICITEM_T1detail3Fields = ICITEM_T1detail3.Fields;
                ICITEM_T1detail4.Compose(new ACCPAC.Advantage.View[] { ICITEM_T1header, null, ICITEM_T1detail1 });
                ICITEM_T1detail4Fields = ICITEM_T1detail4.Fields;
                //    ICITEM_T1detail5.Compose(new ACCPAC.Advantage.View[] { ICITEM_T1header });
                //ICITEM_T1detail5Fields = ICITEM_T1detail5.Fields;
                ICITEM_T1detail6.Compose(new ACCPAC.Advantage.View[] { ICITEM_T1header });
                ICITEM_T1detail6Fields = ICITEM_T1detail6.Fields;
                //    ICITEM_T1detail7.Compose(new ACCPAC.Advantage.View[] { ICITEM_T1header, null });
                //    ICITEM_T1detail7Fields = ICITEM_T1detail7.Fields;
                ICITEM_T1detail8.Compose(new ACCPAC.Advantage.View[] { ICITEM_T1header });
                ICITEM_T1detail8Fields = ICITEM_T1detail10.Fields;
                ICITEM_T1detail9.Compose(new ACCPAC.Advantage.View[] { ICITEM_T1header, null, ICITEM_T1detail1 });
                ICITEM_T1detail9Fields = ICITEM_T1detail9.Fields;
                ICITEM_T1detail10.Compose(new ACCPAC.Advantage.View[] { ICITEM_T1header });
                ICITEM_T1detail10Fields = ICITEM_T1detail10.Fields;
                ICITEM_T1detail11.Compose(new ACCPAC.Advantage.View[] { ICITEM_T1header });
                ICITEM_T1detail11Fields = ICITEM_T1detail11.Fields;

                ICILOC1 = mDBLinkCmpRW.OpenView("IC0290");
                ICILOC1Fields = ICILOC1.Fields;

                return (true);
            }
            catch (Exception ex)
            {
                string message = "Cannot open ICITEM views in OpenAndComposeItemViews()";

                FireEvent("Cannot open ICITEM views in OpenAndComposeItemViews(");
                MyErrorHandler(ex, message);
                return (false);
            }
        }

        public void CleanUpItemViews()
        {
            try
            {
                ICITEM_T1headerFields.Dispose();
                ICITEM_T1detail1Fields.Dispose();
                ICITEM_T1detail2Fields.Dispose();
                ICITEM_T1detail3Fields.Dispose();
                ICITEM_T1detail4Fields.Dispose();
                //ICITEM_T1detail5Fields.Dispose();
                ICITEM_T1detail6Fields.Dispose();
                //ICITEM_T1detail7Fields.Dispose();
                ICITEM_T1detail8Fields.Dispose();
                ICITEM_T1detail9Fields.Dispose();
                ICITEM_T1detail10Fields.Dispose();
                ICITEM_T1detail11Fields.Dispose();
                ICITEM_T1header.Dispose();
                ICITEM_T1detail1.Dispose();
                ICITEM_T1detail2.Dispose();
                ICITEM_T1detail3.Dispose();
                ICITEM_T1detail4.Dispose();
                //ICITEM_T1detail5.Dispose();
                ICITEM_T1detail6.Dispose();
                //ICITEM_T1detail7.Dispose();
                ICITEM_T1detail8.Dispose();
                ICITEM_T1detail9.Dispose();
                ICITEM_T1detail10.Dispose();
                ICITEM_T1detail11.Dispose();
            }
            catch (Exception ex)
            {
                string message = "Cannot dispose of ICITEM views in CleanupItemViews()";
                FireEvent("Cannot dispose of ICITEM views in CleanupItemViews()");
                MyErrorHandler(ex, message);
            }
        }

        private void FireEvent(string message)
        {
            MyEvent e1 = new MyEvent();
            e1.message = message;

            if (Event1 != null)
            {
                Event1(this, e1);
            }

            e1 = null;
        }

        //
        // Simple generic error handler.
        //

        private void MyErrorHandler(Exception e, string message)
        {
            long lCount;
            int iIndex;

            if (sageSession.Errors == null)
            {
                if (Program.sMode != "Silent")
                    MessageBox.Show(message);
                else
                    email.Sendemail("PDA Bad Import", message, "", Program.emailBadData, Program.emailBadDatacc, "", arrayofSalesReps);
            }
            else
            {
                lCount = sageSession.Errors.Count;

                if (lCount == 0)
                {
                    if (Program.sMode != "Silent")
                        MessageBox.Show(message + Environment.NewLine + e.Message);
                }
                else
                {
                    for (iIndex = 0; iIndex < lCount; iIndex++)
                    {
                        message = message + Environment.NewLine + e.Message;
                    }

                    if (Program.sMode != "Silent")
                        MessageBox.Show(message);
                    else
                        email.Sendemail("PDA Bad Import", message, "", System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"], "", arrayofSalesReps);

                    sageSession.Errors.Clear();
                }
            }

        }
    }
}

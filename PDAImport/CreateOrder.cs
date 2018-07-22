using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using ACCPAC.Advantage;
using System.Text.RegularExpressions;

namespace PDAImport
{
    class CreateOrder
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
        private Session session;
        private DBLink mDBLinkCmpRW;

        private ACCPAC.Advantage.View OEORD1header;
        private ACCPAC.Advantage.View OEORD1headerFields;
        private ACCPAC.Advantage.View OEORD1detail1;
        private ACCPAC.Advantage.View OEORD1detail1Fields;
        private ACCPAC.Advantage.View OEORD1detail2;
        private ACCPAC.Advantage.View OEORD1detail2Fields;
        private ACCPAC.Advantage.View OEORD1detail3;
        private ACCPAC.Advantage.View OEORD1detail3Fields;
        private ACCPAC.Advantage.View OEORD1detail4;
        private ACCPAC.Advantage.View OEORD1detail4Fields;
        private ACCPAC.Advantage.View OEORD1detail5;
        private ACCPAC.Advantage.View OEORD1detail5Fields;
        private ACCPAC.Advantage.View OEORD1detail6;
        private ACCPAC.Advantage.View OEORD1detail6Fields;
        private ACCPAC.Advantage.View OEORD1detail7;
        private ACCPAC.Advantage.View OEORD1detail7Fields;
        private ACCPAC.Advantage.View OEORD1detail8;
        private ACCPAC.Advantage.View OEORD1detail8Fields;
        private ACCPAC.Advantage.View OEORD1detail9;
        private ACCPAC.Advantage.View OEORD1detail9Fields;
        private ACCPAC.Advantage.View OEORD1detail10;
        private ACCPAC.Advantage.View OEORD1detail10Fields;
        private ACCPAC.Advantage.View OEORD1detail11;
        private ACCPAC.Advantage.View OEORD1detail11Fields;
        private ACCPAC.Advantage.View OEORD1detail12;
        private ACCPAC.Advantage.View OEORD1detail12Fields;

        private ACCPAC.Advantage.View OESHI1header;
        private ACCPAC.Advantage.View OESHI1headerFields;
        private ACCPAC.Advantage.View OESHI1detail1;
        private ACCPAC.Advantage.View OESHI1detail1Fields;
        private ACCPAC.Advantage.View OESHI1detail2;
        private ACCPAC.Advantage.View OESHI1detail2Fields;
        private ACCPAC.Advantage.View OESHI1detail3;
        private ACCPAC.Advantage.View OESHI1detail3Fields;
        private ACCPAC.Advantage.View OESHI1detail4;
        private ACCPAC.Advantage.View OESHI1detail4Fields;
        private ACCPAC.Advantage.View OESHI1detail5;
        private ACCPAC.Advantage.View OESHI1detail5Fields;
        private ACCPAC.Advantage.View OESHI1detail6;
        private ACCPAC.Advantage.View OESHI1detail6Fields;
        private ACCPAC.Advantage.View OESHI1detail7;
        private ACCPAC.Advantage.View OESHI1detail7Fields;
        private ACCPAC.Advantage.View OESHI1detail8;
        private ACCPAC.Advantage.View OESHI1detail8Fields;
        private ACCPAC.Advantage.View OESHI1detail9;
        private ACCPAC.Advantage.View OESHI1detail9Fields;
        private ACCPAC.Advantage.View OESHI1detail10;
        private ACCPAC.Advantage.View OESHI1detail10Fields;
        private ACCPAC.Advantage.View OESHI1detail11;
        private ACCPAC.Advantage.View OESHI1detail11Fields;
        private ACCPAC.Advantage.View OESHI1detail12;
        private ACCPAC.Advantage.View OESHI1detail12Fields;

        private ACCPAC.Advantage.View ICITEM_T1header;
        private ACCPAC.Advantage.View ICITEM_T1headerFields;
        private ACCPAC.Advantage.View ICITEM_T1detail1;
        private ACCPAC.Advantage.View ICITEM_T1detail1Fields;
        private ACCPAC.Advantage.View ICITEM_T1detail2;
        private ACCPAC.Advantage.View ICITEM_T1detail2Fields;
        private ACCPAC.Advantage.View ICITEM_T1detail3;
        private ACCPAC.Advantage.View ICITEM_T1detail3Fields;
        private ACCPAC.Advantage.View ICITEM_T1detail4;
        private ACCPAC.Advantage.View ICITEM_T1detail4Fields;
        private ACCPAC.Advantage.View ICITEM_T1detail5;
        private ACCPAC.Advantage.View ICITEM_T1detail5Fields;
        private ACCPAC.Advantage.View ICITEM_T1detail6;
        private ACCPAC.Advantage.View ICITEM_T1detail6Fields;
        private ACCPAC.Advantage.View ICITEM_T1detail7;
        private ACCPAC.Advantage.View ICITEM_T1detail7Fields;
        private ACCPAC.Advantage.View ICITEM_T1detail8;
        private ACCPAC.Advantage.View ICITEM_T1detail8Fields;

        private bool PostOrder;
        private bool PostShipment;
        private bool PrintShipment;
        private bool PrintConfirmation;
        private string CustomerEmail;
        private float TotalQtyShipped;

        // Global Variables
        public string sSourcePath;
        public string sBackupPath;
        public string sErrorPath;
        public int dateFormat;
        public string UserName;
        public object txtwrite;
        public string txtInputFile;
        public object txtBadData;
        public object txtOnHold;
        public object txtNoQty;
        public string sWarehouse;
        public string sExclude;
        public string txtOutputFile;
        public string txtOnHoldFile;
        public string txtNoQtyFile;
        public string txtBadDataFile;
        public string txtOrderType;

        public bool SendErrorEmail;
        public string txtTOROnHoldEmail;
        public string txtMTLOnHoldEmail;
        public string txtOnHoldEmailcc;
        public string txtBadDataEmail;
        public string txtBadDataEmailcc;
        public string txtShortageEmail;
        public string txtShortageEmailcc;
        public string txtNoQtyEmail;
        public string txtNoQtyEmailcc;
        public string txtTORAROverDueEmail;
        public string txtMTLAROverDueEmail;
        public string txtAROverDueEmailcc;
        public string sTest;
        public string sMove;
        public int iTmp;
        public string ItemDescription;
        public int iDaysOverDue;

        public string[] SalesReps;
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
        }

        public Boolean PerformPDAImport(string sLoc)
        {

            LogWriter log = new LogWriter();

            List<string> ListofFiles = new List<string>();
            ListofFiles = GetFileNames();

            if (ListofFiles.Count > 0)
                {

                if (Program.sMode != "silent")
                {
                    FireEvent("Logging into Sage");
                }

                if (SageLogin())
                {
                    log.LogWrite(Program.backupPath, Program.txtOutputFile, "Successfully logged into Sage.");
                    FireEvent("Successfully logged in.");

                    arrayofSalesReps = GetSalesReps();

                    foreach (string fullFileName in ListofFiles)
                    {
                        FireEvent("Processing file " + fullFileName);

                        /* https://stackoverflow.com/questions/2081418/parsing-csv-files-in-c-with-header
                           Using the generic ReadCSVFile code at the end */

                        string status = ProcessFile(fullFileName, sLoc);
                    }
                    return true;
                }
                else
                {
                    FireEvent("log into Sage failed.");
                    log.LogWrite(Program.backupPath, Program.txtOutputFile, "Log into Sage failed.");
                    return false;
                }
            }
            FireEvent("No files to process.");
            log.LogWrite(Program.backupPath, Program.txtOutputFile, "No files to process.");
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
                session = new Session();
                session.Init("", "XX", "XX1000", "63A");
                session.Open("ADMIN", "RTFM", "JFJDAT", DateTime.Today, 0);

                // Open a database link.
                mDBLinkCmpRW = session.OpenDBLink(DBLinkType.Company, DBLinkFlags.ReadWrite);

                ACCPAC.Advantage.FiscalCalendar fiscCal = mDBLinkCmpRW.FiscalCalendar;

                bRet = fiscCal.GetPeriod(new DateTime(DateTime.Now.Year,DateTime.Now.Month,DateTime.Now.Day), out periods, out year, out bActive);
                return true;
            }
            catch (Exception ex)
            {
                string message = "Cannot log into Sage";
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
                if (Program.sLoc == "VAN")
                {
                    ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"D:\KMS\PDA\VA\var\spool\jetsso\SFA\", "*", SearchOption.AllDirectories));
                    ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"D:\KMS\PDA\VA\var\spool\jetsso\PDA\", "*", SearchOption.AllDirectories));
                    ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"D:\PDA_Files\VA\REPROCESS\", "*", SearchOption.AllDirectories));
                }
                else
                {
                    ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"D:\KMS\PDA\TR\var\spool\jetsso\SFA\", "*", SearchOption.AllDirectories));
                    ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"D:\KMS\PDA\TR\var\spool\jetsso\PDA\", "*", SearchOption.AllDirectories));
                    ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"D:\PDA_Files\TR\REPROCESS\", "*", SearchOption.AllDirectories));
                }
            }
            else
            {
                if (Program.sLoc == "VAN")
                {
                    ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"c:\Files\Clients\JFC\PDA Files\VAN\SFA\", "*", SearchOption.AllDirectories));
                    ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"c:\Files\Clients\JFC\PDA Files\VAN\PDA\", "*", SearchOption.AllDirectories));
                    ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"c:\Files\Clients\JFC\PDA Files\REPROCESS", "*", SearchOption.AllDirectories));
                }
                else
                {
                    ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"c:\Files\Clients\JFC\PDA Files\TOR\SFA\", "*", SearchOption.AllDirectories));
                    ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"c:\Files\Clients\JFC\PDA Files\TOR\PDA\", "*", SearchOption.AllDirectories));
                    ListofFiles = FilterFiles(ListofFiles, Directory.GetFiles(@"c:\Files\Clients\JFC\PDA Files\REPROCESS\", "*", SearchOption.AllDirectories));
                }
            }

            return ListofFiles;
        }

        private List<string> FilterFiles(List<string> ListofFiles, string[] ArrayofFiles)
        {
            FileInfo fileInfo;
            Regex reg = new Regex(@"^[frs]\d{9,10}\.\d{4}$");

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

            string[] lines = File.ReadAllLines(salesRepsPath);
            string[][] locArrayofSalesReps = lines.Select(line => line.Split(',').ToArray()).ToArray();

            return locArrayofSalesReps;
        }

//        string ProcessFile(string[] arrayofLines)
        string ProcessFile(string fullFileName, string sLoc)
        {
            string status;
            string strLine;
            string sRowType;
            string shipmentStatus;
            long lastLine;
            string uniqueNumber;
            string reference;
            DateTime ordDate;
            DateTime expDate;
            DateTime requestDate;
            string txtYr;
            string txtMth;
            string txtDay;
            string sFolder;
            int iCnt;
            long iNumRows;
            int iNumColumns;
            int iUniqueNumIndex;
            int iCustomerIndex;
            int iReferenceIndex;
            int iOrderDateIndex;
            int iExpDateIndex;
            int iLocationIndex;
            int iCommentIndex;
            int iPONumberIndex;
            int iShipViaIndex;
            int iShipToIndex;
            int iLineNumIndex;
            int iLineTypeIndex;
            int iItemIndex;
            int iMiscChargeIndex;
            int iQtyOrderedIndex;
            int iOrdUnitIndex;
            int iPriUnitPrcIndex;
            int iExtInvMiscIndex;
            bool stillOk;
            string NewOrderNumber;
            string newShipmentNumber;
            string ItemNo;
            float qtyOrdered;
            float qtyShipped;
            string Location = string.Empty;
            string OrderUnit;
            string poNumber = string.Empty;
            string comment = string.Empty;
            string shipVia = string.Empty;
            string shipTo = string.Empty;
            string ItemStatus;
            string[] sSalesPerson;
            string custNo = string.Empty;
            string sendeMailStatus;
            string sSubject;
            string custName;
            string message;
            string[] ItemsShippedItemno;
            string[] ItemsShippedQty;
            int iSalesRep;
            string sSalesRepLoc;
            string stringToBeFound;
            string SaveBodyText;
            bool shortShipped;
            bool postShipment;
            bool printShipment;
            bool postOrder;
            bool printConfirmation;
            float totalQtyShipped;

            status = OpenAndComposeOrderViews();

            iNumRows = 0;
            sRowType = "H";
            stillOk = true;
            shipmentStatus = "OK";
            postOrder = true;
            postShipment = true;
            printShipment = true;
            printConfirmation = false;
            shortShipped = false;
            message = "";
            totalQtyShipped = 0;
            iUniqueNumIndex = -1;
            iCustomerIndex = -1;
            iReferenceIndex = -1;
            iOrderDateIndex = -1;
            iExpDateIndex = -1;
            iLocationIndex = -1;
            iCommentIndex = -1;
            iPONumberIndex = -1;
            iShipViaIndex = -1;
            iShipToIndex = -1;
            iLineNumIndex = -1;
            iLineTypeIndex = -1;
            iItemIndex = -1;
            iMiscChargeIndex = -1;
            iQtyOrderedIndex = -1;
            iOrdUnitIndex = -1;
            iPriUnitPrcIndex = -1;
            iExtInvMiscIndex = -1;
            sendeMailStatus = "NO";
            custName = "";
            sSubject = "";
            sSalesRepLoc = string.Empty;
            reference = string.Empty;

            List<PDAImport.ReadCSV.CSV.Record> list = PDAImport.ReadCSV.CSV.ParseFile(fullFileName);

            for (var irow = 0; irow < list.Count; irow++)
            {
                if (irow == 0)
                {
                    for (var icol = 0; icol < list[irow].Row.Length; icol++)
                    {
                        switch (list[irow].Row[icol].ToString())
                        {
                            case "ORDUNIQ":
                                iUniqueNumIndex = icol;
                                break;
                            case "CUSTOMER":
                                iCustomerIndex = icol;
                                break;
                            case "REFERENCE":
                                iReferenceIndex = icol;
                                break;
                            case "ORDDATE":
                                iOrderDateIndex = icol;
                                break;
                            case "EXPDATE":
                                iExpDateIndex = icol;
                                break;
                            case "LOCATION":
                                iLocationIndex = icol;
                                break;
                            case "PONUMBER":
                                iPONumberIndex = icol;
                                break;
                            case "SHIPVIA":
                                iShipViaIndex = icol;
                                break;
                            case "COMMENT":
                                iCommentIndex = icol;
                                break;
                            case "SHIPTO":
                                iShipToIndex = icol;
                                break;
                        }
                    }
                }
                else if (irow == 1)
                {
                    uniqueNumber = list[irow].Row[iUniqueNumIndex].ToString();
                    custNo = list[irow].Row[iCustomerIndex].ToString();
                    reference = list[irow].Row[iReferenceIndex].ToString();
                    sSalesPerson = reference.Split('-');
                    ordDate = DateTime.ParseExact(list[irow].Row[iOrderDateIndex].ToString(), "yyyyMMdd", null);
                    expDate = DateTime.ParseExact(list[irow].Row[iOrderDateIndex].ToString(), "yyyyMMdd", null);
                    requestDate = expDate;
                    stringToBeFound = sSalesPerson[0];
                    iSalesRep = IsInArray(stringToBeFound, arrayofSalesReps);
                    if (iSalesRep >= 0)
                        sSalesRepLoc = arrayofSalesReps[iSalesRep][2];
                    else
                        sSalesRepLoc = "TOR";

                    if (sSalesRepLoc == "TOR" & !(sLoc == "TOR" | sLoc =="TORMTL"))
                    {
                        CleanUpOrderViews();
                        return ("SKIP");
                    }

                    if (sSalesRepLoc == "MTL" & !(sLoc == "MTL" | sLoc == "TORMTL"))
                    {
                        CleanUpOrderViews();
                        return ("SKIP");
                    }

                    if (sSalesRepLoc == "VAN" & !(sLoc == "VAN" | sLoc == "VANCAL"))
                    {
                        CleanUpOrderViews();
                        return ("SKIP");
                    }

                    if (sSalesRepLoc == "CAL" & !(sLoc == "CAL" | sLoc == "VANCAL"))
                    {
                        CleanUpOrderViews();
                        return ("SKIP");
                    }

                    if (iLocationIndex >= 0)
                        Location = list[irow].Row[iLocationIndex].ToString();

                    if (Location != "1" & sLoc == "TOR")
                    {
                        CleanUpOrderViews();
                        return ("SKIP");
                    }

                    if (Location != "2" & sLoc == "VAN")
                    {
                        CleanUpOrderViews();
                        return ("SKIP");
                    }

                    if (iPONumberIndex >= 0)
                        poNumber = list[irow].Row[iPONumberIndex].ToString();

                    if (iCommentIndex >= 0)
                        comment = list[irow].Row[iCommentIndex].ToString();

                    if (iShipViaIndex >= 0)
                        shipVia = list[irow].Row[iShipViaIndex].ToString();

                    if (iShipToIndex >= 0)
                        shipTo = list[irow].Row[iShipToIndex].ToString();

                    status = CheckSalesPerson(sSalesPerson[0], custNo, reference, fullFileName, sendeMailStatus);

                    if (status == "BAD")
                    {
                        CleanUpOrderViews();

                        if (Program.Move == "Yes")
                        {
                            Utilities.MoveFile(sSourcePath, txtInputFile, sErrorPath, Program.sourceSystem, txtBadDataEmail, txtBadDataEmailcc);
                        }
                        return ("BAD");
                    }
                }

            }
            return "OK";
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
                        message = message + "\r\n" + "File - " + fullFileName;
                        message = message + "\r\n" + "SalesPerson - " + sSalesPerson;
                        message = message + "\r\n" + "SalesPerson is inactive.";
                        message = message + "\r\n" + "Reference - " + reference;
                        message = message + "\r\n" + "Customer - " + custNo;
                        email.Sendemail("PDA Bad SalesPerson", message, fullFileName, Program.emailBadData, Program.emailBadDatacc, sSalesPerson);

                        return ("BAD");
                    }
                    else
                        return "OK";
                }
                else
                    message = "PDA Bad Data Import Error";
                    message = message + "\r\n" + "File - " + fullFileName;
                    message = message + "\r\n" + "SalesPerson - " + sSalesPerson;
                    message = message + "\r\n" + "SalesPerson is missing.";
                    message = message + "\r\n" + "Reference - " + reference;
                    message = message + "\r\n" + "Customer - " + custNo;
                    email.Sendemail("PDA Bad SalesPerson", message, fullFileName, Program.emailBadData, Program.emailBadDatacc, sSalesPerson);
                    return "BAD";
            }
            catch (Exception ex)
            {
                message = "PDA Bad Data Import Error";
                message = message + "\r\n" + "File - " + fullFileName;
                message = message + "\r\n" + "SalesPerson - " + sSalesPerson;
                message = message + "\r\n" + ex.Message;
                message = message + "\r\n" + "Reference - " + reference;
                message = message + "\r\n" + "Customer - " + custNo;
                MyErrorHandler(ex, message);
                return ("BAD");
            }

        }

        private string CheckCustomer(string custNo, string reference, string custName, bool printShipment, bool printConfirmation, string sendeMailStatus)
        {
            Boolean fnd;
            string message;
            string status = String.Empty;

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

                    if (Convert.ToInt32(arCustomerHeader.Fields.FieldByName("SWHOD").Value) == 0)
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
                        sSQL = "SELECT [IDCUST] ";
                        sSQL = sSQL + ",[IDINVC] ";
                        sSQL = sSQL + ",[TRXTYPETXT] ";
                        sSQL = sSQL + ",[DATEINVC] ";
                        sSQL = sSQL + ",substring(cast([DATEINVC] as CHAR(8)),1,4) As YEAR ";
                        sSQL = sSQL + ",substring(cast([DATEINVC] as CHAR(8)),5,2) As MONTH ";
                        sSQL = sSQL + ",substring(cast([DATEINVC] as CHAR(8)),7,2) As DAY ";
                        sSQL = sSQL + ",cast(substring(cast([DATEINVC] as CHAR(8)),1,4) + substring(cast([DATEINVC] as CHAR(8)),5,2) + substring(cast([DATEINVC] as CHAR(8)),7,2) AS DATE) AS DUEINVC ";
                    //  sSQL = sSQL + ",DATEDIFF(DAY,GETDATE(),cast(substring(cast([DATEINVC] as CHAR(8)),1,4) + substring(cast([DATEINVC] as CHAR(8)),5,2) + substring(cast([DATEINVC] as CHAR(8)),7,2) AS DATE)) AS DAYS ";
                        sSQL = sSQL + ",DATEDIFF(DAY,cast(substring(cast([DATEINVC] as CHAR(8)),1,4) + substring(cast([DATEINVC] as CHAR(8)),5,2) + substring(cast([DATEINVC] as CHAR(8)),7,2) AS DATE),GETDATE()) AS DAYS ";
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

                return (status);
            }

            catch (Exception ex)
            {
                message = "PDA Bad Data Import Error";
                message = message + "\r\n" + ex.Message;
                message = message + "\r\n" + "Reference - " + reference;
                message = message + "\r\n" + "Customer - " + custNo;
                MyErrorHandler(ex, message);
                return ("BAD");
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

        private int IsInArray(string stringToBeFound, string[][] arr)
        {
            for (int i = 0; i < arr.GetLength(0); i++)
                if (stringToBeFound.Equals(arr[i][0]))
                {
                    return i;
                }
            return -1;
        }

        private void MyErrorHandler(Exception e, string message)
        {
            long lCount;
            int iIndex;

            if (session.Errors == null)
            {
                if (Program.sMode != "silent")
                    MessageBox.Show(e.Message);
                else
                    email.Sendemail("PDA Bad Import", message, "", Program.emailBadData, Program.emailBadDatacc, "");
            }
            else
            {
                lCount = session.Errors.Count;

                if (lCount == 0)
                {
                    if (Program.sMode != "silent")
                        MessageBox.Show(e.Message);
                }
                else
                {
                    for (iIndex = 0; iIndex < lCount; iIndex++)
                    {
                        message = message + "\r\n" + e.Message;
                    }

                    if (Program.sMode != "silent")
                        MessageBox.Show(e.Message);
                    else
                        email.Sendemail("PDA Bad Import", message, "", Program.emailBadData, Program.emailBadDatacc, "");

                    session.Errors.Clear();
                }
            }

        }
    }
}

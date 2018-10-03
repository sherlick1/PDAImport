using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDAImport
{
    public partial class CreateOrder
    {

        public string ProcessFile(string fullFileName, string sLoc)
        {
            string status;
            string uniqueNumber;
            string reference;
            DateTime ordDate = DateTime.Now;
            DateTime expDate = DateTime.Now;
            DateTime requestDate = DateTime.Now;
            long iNumRows;
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
            string newOrderNumber = string.Empty;
            string newShipmentNumber;
            string itemNo = String.Empty;
            string miscCharge = String.Empty;
            double qtyOrdered = 0;
            double qtyShipped = 0;
            double totalQtyShipped = 0;
            decimal priUntPrc = 0;
            decimal extInvMisc = 0;
            string location = string.Empty;
            string orderUnit = string.Empty;
            string poNumber = string.Empty;
            string comment = string.Empty;
            string shipVia = string.Empty;
            string shipTo = string.Empty;
            string itemStatus = string.Empty;
            string[] sSalesPerson = new string[] { };
            string custNo = string.Empty;
            string sendeMailStatus;
            string sSubject;
            string custName;
            string message;
            string bodyText = String.Empty;
            List<ItemsShippedLine> itemsShipped = new List<ItemsShippedLine>();
            int iSalesRep;
            string sSalesRepLoc;
            string stringToBeFound;
            string saveBodyText = String.Empty;
            bool shortShipped;
            bool postShipment;
            bool printShipment;
            bool postOrder;
            bool printConfirmation;
            string lineType = string.Empty;

            string emailBadData = string.Empty;
            string emailBadDatacc = string.Empty;
            string txtEmail = string.Empty;
            string txtEmailcc = string.Empty;

            LogWriter log = new LogWriter();

            status = OpenAndComposeOrderViews();

            iNumRows = 0;
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
                    iSalesRep = Utilities.IsInArray(stringToBeFound, arrayofSalesReps);
                    if (iSalesRep >= 0)
                        sSalesRepLoc = arrayofSalesReps[iSalesRep][2];
                    else
                    {
                        if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                        {
                            emailBadData = System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data"];
                            emailBadDatacc = System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data_cc"];
                        }
                        else
                        {
                            emailBadData = System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"];
                            emailBadDatacc = System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"];
                        }
                        email.Sendemail("PDA macro Sales Rep missing from Salesrep.ini file", "Sales Rep " + sSalesPerson[0] + " is missing from the Salesrep.ini file", null, emailBadData, emailBadDatacc, null, null);
                        CleanUpOrderViews();
                        return ("SKIP");
                    }

                    if (sSalesRepLoc == "TOR" & !(sLoc == "TOR" | sLoc == "TORMTL"))
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
                        location = list[irow].Row[iLocationIndex].ToString();

                    if (location != "1" & sLoc == "TOR")
                    {
                        CleanUpOrderViews();
                        return ("SKIP");
                    }

                    if (location != "2" & sLoc == "VAN")
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

                    status = CheckSalesPerson(sSalesPerson[0], custNo, fullFileName, fullFileName, sendeMailStatus);

                    if (status == "BAD")
                    {
                        CleanUpOrderViews();

                        if (Program.Move == "Yes")
                        {
                            if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                                Utilities.MoveFile(fullFileName, Program.torbackupPath, System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"]);

                            if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                                Utilities.MoveFile(fullFileName, Program.vanbackupPath, System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data_cc"]);

                        }
                        return ("BAD");
                    }

                    status = CheckCustomer(custNo, fullFileName, ref custName, printShipment, printConfirmation, sendeMailStatus);

                    if (status == "BAD")
                    {
                        CleanUpOrderViews();

                        if (Program.Move == "Y")
                        {
                            if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                                Utilities.MoveFile(fullFileName, System.Configuration.ConfigurationManager.AppSettings["tor_error_path"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"]);

                            if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                                Utilities.MoveFile(fullFileName, System.Configuration.ConfigurationManager.AppSettings["van_error_path"], System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data_cc"]);
                        }
                        return ("BAD");
                    }

                    if (!(String.IsNullOrEmpty(shipTo)))
                    {
                        status = CheckShipToCustomer(custNo, fullFileName, custName, shipTo, sendeMailStatus);
                    }

                    if (status == "BAD")
                    {
                        CleanUpOrderViews();

                        if (Program.Move == "Y")
                        {
                            if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                                Utilities.MoveFile(fullFileName, System.Configuration.ConfigurationManager.AppSettings["tor_error_path"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"]);

                            if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                                Utilities.MoveFile(fullFileName, System.Configuration.ConfigurationManager.AppSettings["van_error_path"], System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data_cc"]);
                        }
                        return ("BAD");
                    }

                }
                else if (irow == 2)
                {
                    try
                    {
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

                        OEORD1header.Cancel();
                        OEORD1header.Init();
                        // Object customerCodeObject = (Object)custNo;
                        // oeOrdHeader.Fields.FieldByName("CUSTOMER").SetValue(customerCodeObject, false);

                        OEORD1headerFields.FieldByName("CUSTOMER").SetValue(custNo, false);

                        if (!(String.IsNullOrEmpty(shipTo)))
                        {
                            OEORD1headerFields.FieldByName("SHIPTO").SetValue(shipTo, false);          // Ship to location
                        }

                        OEORD1headerFields.FieldByName("PONUMBER").SetValue(poNumber, false);          // Purchase Order Number
                        OEORD1headerFields.FieldByName("ORDDATE").SetValue(ordDate, false);            // Order Date
                        OEORD1headerFields.FieldByName("EXPDATE").SetValue(expDate, false);            // Expected Date
                        OEORD1headerFields.FieldByName("LOCATION").SetValue(location, false);          // Default Location Code
                                                                                                       //  OEORD1header.Fields.FieldByName("ONHOLD").PutWithoutVerification ("0")         // On Hold Status
                        OEORD1headerFields.FieldByName("SHIPVIA").SetValue(shipVia, false);            // Ship-Via Code
                        OEORD1headerFields.FieldByName("REQUESDATE").SetValue(requestDate, false);     // Requested Date
                        OEORD1headerFields.FieldByName("REFERENCE").SetValue(reference, false);        // SalesPerson with hyphen and number
                        OEORD1headerFields.FieldByName("COMMENT").SetValue(comment, false);            // Comment
                        OEORD1headerFields.FieldByName("SALESPER1").SetValue(sSalesPerson[0], false);  // SalesPerson - Use Default Sales Rep off of customer master record
                        OEORD1headerFields.FieldByName("ONHOLD").SetValue(false, false);               // Do not put order on hold

                    }

                    catch (Exception ex)
                    {
                        message = "PDA Bad Data Import Error";
                        message = message + Environment.NewLine + "File - " + fullFileName;
                        message = message + Environment.NewLine + "Customer - " + custNo;
                        message = message + Environment.NewLine + ex.Message;
                        message = message + Environment.NewLine + "Reference - " + reference;
                        MyErrorHandler(ex, message);
                        return ("BAD");
                    }

                    // Find Detail Line Columns
                    for (var icol = 0; icol < list[irow].Row.Length; icol++)
                    {
                        switch (list[irow].Row[icol].ToString())
                        {
                            case "ORDUNIQ":
                                iUniqueNumIndex = icol;
                                break;
                            case "LINENUM":
                                iLineNumIndex = icol;
                                break;
                            case "LINETYPE":
                                iLineTypeIndex = icol;
                                break;
                            case "ITEM":
                                iItemIndex = icol;
                                break;
                            case "MISCCHARGE":
                                iMiscChargeIndex = icol;
                                break;
                            case "QTYORDERED":
                                iQtyOrderedIndex = icol;
                                break;
                            case "ORDUNIT":
                                iOrdUnitIndex = icol;
                                break;
                            case "PRIUNTPRC":
                                iPriUnitPrcIndex = icol;
                                break;
                            case "EXTINVMISC":
                                iExtInvMiscIndex = icol;
                                break;
                        }
                    }
                }
                else if (irow >= 3)
                {
                    OEORD1headerFields.FieldByName("PROCESSCMD").SetValue("1", false);     // Process OIP Command
                    OEORD1header.Process();

                    OEORD1detail1.Read(false);
                    OEORD1detail1.RecordCreate(0);
                    OEORD1detail1Fields.FieldByName("LINENUM").SetValue(Convert.ToString((iNumRows + 1) * -1), false);

                    if (iLineTypeIndex > 0)
                        lineType = list[irow].Row[iLineTypeIndex].ToString();
                    else
                        lineType = "1"; // Set to "Item" type if no lone type in the input file

                    if (iItemIndex > 0)
                        itemNo = list[irow].Row[iItemIndex].ToString();
                    else
                        itemNo = "";

                    if (iMiscChargeIndex > 0)
                        miscCharge = list[irow].Row[iMiscChargeIndex].ToString();
                    else
                        miscCharge = "";

                    if (iQtyOrderedIndex > 0)
                        qtyOrdered = Convert.ToDouble(list[irow].Row[iQtyOrderedIndex]);
                    else
                        qtyOrdered = 0;

                    if (iOrdUnitIndex > 0)
                        orderUnit = list[irow].Row[iOrdUnitIndex].ToString();
                    else
                        orderUnit = "";

                    if (iPriUnitPrcIndex > 0)
                        priUntPrc = Convert.ToDecimal(list[irow].Row[iPriUnitPrcIndex]);
                    else
                        priUntPrc = 0;

                    if (iExtInvMiscIndex > 0)
                        extInvMisc = Convert.ToDecimal(list[irow].Row[iExtInvMiscIndex]);
                    else
                        extInvMisc = 0;

                    if (lineType == "1")
                    {
                        itemStatus = "OK";
                        qtyShipped = 0;
                        CheckItem(itemNo, fullFileName, location, qtyOrdered, orderUnit, ref itemStatus, postOrder, ref qtyShipped, itemsShipped, ref priUntPrc);

                        if (itemStatus == "OK" || itemStatus == "Not_Enough_Quantity")
                        {
                            OEORD1detail1Fields.FieldByName("ITEM").SetValue(itemNo, false);               // Item
                            OEORD1detail1Fields.FieldByName("QTYORDERED").SetValue(qtyOrdered, false);     // Quantity Ordered
                            if (postShipment)
                            {
                                OEORD1detail1Fields.FieldByName("QTYSHIPPED").SetValue(qtyShipped, false); // Quantity Shipped
                                totalQtyShipped = totalQtyShipped + qtyShipped;
                            }

                            if (itemStatus == "Not_Enough_Quantity")
                                shortShipped = true;

                            try
                            {
                                OEORD1detail1Fields.FieldByName("PROCESSCMD").SetValue("1", false);        // Process Command
                                OEORD1detail1.Process();

                                OEORD1detail1Fields.FieldByName("ORDUNIT").SetValue(orderUnit, false);            // Order Unit
                                OEORD1detail1Fields.FieldByName("PRIUNTPRC").SetValue(priUntPrc, false);                 // Pricing Unit Price
                                OEORD1detail1.Insert();
                            }
                            catch (Exception ex)
                            {
                                message = "PDA Bad Data Import Error in CreateOrder - Inserting Item.";
                                message = message + Environment.NewLine + "File - " + fullFileName;
                                message = message + Environment.NewLine + "SalesPerson - " + sSalesPerson[0].ToString();
                                message = message + Environment.NewLine + "Reference - " + reference;
                                message = message + Environment.NewLine + "Customer - " + custNo;
                                message = message + Environment.NewLine + "Item - " + itemNo;
                                message = message + Environment.NewLine + "Location - " + location;
                                message = message + Environment.NewLine + "Order Unit - " + orderUnit;
                                message = message + Environment.NewLine + "Pricing Unit Price - " + priUntPrc;
                                MyErrorHandler(ex, message);
                                return ("BAD");
                            }
                        }
                    }
                    else
                    {
                        OEORD1detail1Fields.FieldByName("LINETYPE").SetValue("1", false);
                        OEORD1detail1Fields.FieldByName("MISCCHARGE").SetValue("1", false);    // Miscellaneous Charges Code
                        OEORD1detail1Fields.FieldByName("PROCESSCMD").SetValue("1", false);    // Process Command
                        OEORD1detail1.Process();
                        OEORD1detail1Fields.FieldByName("EXTINVMISC").SetValue(extInvMisc, false);
                        OEORD1detail1.Insert();
                    }
                }
            }

            if ((postOrder && postShipment && totalQtyShipped > 0) || (postOrder && postShipment == false))
            {
                try
                {
                    OEORD1headerFields.FieldByName("GOCHKCRDT").SetValue("1", false);    // Perform Credit Limit Check 0 = False, 1 = True

                    OEORD1header.Process();
                }

                catch (Exception ex)
                {
                    if (Convert.ToBoolean(OEORD1headerFields.FieldByName("OVERCREDIT").Value))
                    {
                        OEORD1headerFields.FieldByName("APPROVEBY").SetValue("ADMIN", false);    // Authorizing User ID
                        OEORD1headerFields.FieldByName("APPPASSWRD").SetValue("RTFM", false);    // Authorizing User Password
                        OEORD1headerFields.FieldByName("OECOMMAND").SetValue("6", false);        // Process O/E Command
                        OEORD1header.Process();
                        OEORD1headerFields.FieldByName("OECOMMAND").SetValue("4", false);        // Process O/E Command
                        OEORD1header.Process();
                    }
                }

                if (postOrder)
                {
                    OEORD1header.Insert();
                    if (OEORD1header.LastReturnCode == 0 || OEORD1header.LastReturnCode == 6666)
                    {
                        newOrderNumber = OEORD1headerFields.FieldByName("ORDNUMBER").Value.ToString();
                        //if (printConfirmation)
                        //{
                        //  PrintOEConfirmation(newOrderNumber);
                        //  Sendemail(sSourceSystem, "Order Confirmation from JFC", Message, "c:\temp\OE_Confirmation.pdf", CustomerEmail, txtBadDataEmailcc, sSalesPerson)
                        //                    }    
                        newShipmentNumber = OEORD1headerFields.FieldByName("LASTSHINUMBER").Value.ToString();
                        if (postShipment && (newShipmentNumber != ""))
                        {
                            if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                                log.LogWrite(Program.torbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + location + "," + newOrderNumber + "," + newShipmentNumber + "," + custNo + "," + reference + ",Successfully imported");

                            if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                                log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + location + "," + newOrderNumber + "," + newShipmentNumber + "," + custNo + "," + reference + ",Successfully imported");

                            if (printShipment)
                            {
                                if (Program.sMode != "Silent")
                                {
                                    FireEvent("Printing shipment - " + newShipmentNumber);
                                }
                                Printing.PrintPickingSlip(newShipmentNumber, sLoc, sageSession);
                                //Call UpdateShipment(newShipmentNumber, Status)

                                if (shortShipped)
                                {
                                    message = "The following Shipment was short shipped:" + Environment.NewLine;
                                    message = message + Environment.NewLine + "Order Number - " + newOrderNumber + Environment.NewLine;
                                    message = message + Environment.NewLine + "Shipment Number - " + newShipmentNumber + Environment.NewLine;
                                    message = message + Environment.NewLine + "Reference - " + reference + Environment.NewLine;
                                    message = message + Environment.NewLine + "Customer - " + custNo + " - " + custName + Environment.NewLine;
                                    message = message + Environment.NewLine + "File - " + fullFileName + Environment.NewLine;
                                    message = message + Environment.NewLine + bodyText;
                                    sSubject = "Shortages - " + custNo + " - " + custName + " - " + reference;
                                    if (sSalesRepLoc == "MTL")
                                        email.Sendemail(sSubject, message, null, System.Configuration.ConfigurationManager.AppSettings["tor_email_shortage"], System.Configuration.ConfigurationManager.AppSettings["tor_email_shortage_cc"], sSalesPerson[0], arrayofSalesReps);
                                    else if (sSalesRepLoc == "TOR")
                                        email.Sendemail(sSubject, message, null, System.Configuration.ConfigurationManager.AppSettings["tor_email_shortage"], System.Configuration.ConfigurationManager.AppSettings["tor_email_shortage_cc"], sSalesPerson[0], arrayofSalesReps);
                                    else if (sSalesRepLoc == "VAN")
                                        email.Sendemail(sSubject, message, null, System.Configuration.ConfigurationManager.AppSettings["van_email_shortage"], System.Configuration.ConfigurationManager.AppSettings["van_email_shortage_cc"], sSalesPerson[0], arrayofSalesReps);
                                    else if (sSalesRepLoc == "CAL")
                                        email.Sendemail(sSubject, message, null, System.Configuration.ConfigurationManager.AppSettings["van_email_shortage"], System.Configuration.ConfigurationManager.AppSettings["van_email_shortage_cc"], sSalesPerson[0], arrayofSalesReps);
                                    else
                                        email.Sendemail(sSubject, message, null, System.Configuration.ConfigurationManager.AppSettings["tor_email_shortage"], System.Configuration.ConfigurationManager.AppSettings["tor_email_shortage_cc"], sSalesPerson[0], arrayofSalesReps);
                                }

                                if (status == "OVERDUE")
                                {
                                    sSubject = "OVERDUE INVOICES - " + custNo + " - " + custName + " - " + newOrderNumber + " - " + newShipmentNumber;
                                    bodyText = saveBodyText;
                                    bodyText = bodyText + Environment.NewLine + "Customer has overdue invoices greater than " + Program.DaysOverDue.ToString() + " days." + Environment.NewLine;
                                    bodyText = bodyText + Environment.NewLine + "File - " + fullFileName + Environment.NewLine;
                                    bodyText = bodyText + Environment.NewLine + "Customer - " + custNo + Environment.NewLine;
                                    bodyText = bodyText + Environment.NewLine + "Order Number - " + newOrderNumber + Environment.NewLine;
                                    bodyText = bodyText + Environment.NewLine + "Shipment Number - " + newShipmentNumber + Environment.NewLine;
                                    bodyText = bodyText + Environment.NewLine + "Reference - " + reference + Environment.NewLine;
                                    if (sSalesRepLoc == "MTL")
                                        email.Sendemail(sSubject, bodyText, null, System.Configuration.ConfigurationManager.AppSettings["mtl_email_ar_overdue"], System.Configuration.ConfigurationManager.AppSettings["tor_email_ar_overdue_cc"], sSalesPerson[0], arrayofSalesReps);
                                    else if (sSalesRepLoc == "TOR")
                                        email.Sendemail(sSubject, bodyText, null, System.Configuration.ConfigurationManager.AppSettings["tor_email_ar_overdue"], System.Configuration.ConfigurationManager.AppSettings["tor_email_ar_overdue_cc"], sSalesPerson[0], arrayofSalesReps);
                                    else if (sSalesRepLoc == "VAN")
                                        email.Sendemail(sSubject, bodyText, null, System.Configuration.ConfigurationManager.AppSettings["van_email_ar_overdue"], System.Configuration.ConfigurationManager.AppSettings["van_email_ar_overdue_cc"], sSalesPerson[0], arrayofSalesReps);
                                    else if (sSalesRepLoc == "CAL")
                                        email.Sendemail(sSubject, bodyText, null, System.Configuration.ConfigurationManager.AppSettings["cal_email_ar_overdue"], System.Configuration.ConfigurationManager.AppSettings["van_email_ar_overdue_cc"], sSalesPerson[0], arrayofSalesReps);
                                    else
                                        email.Sendemail(sSubject, bodyText, null, System.Configuration.ConfigurationManager.AppSettings["tor_email_ar_overdue"], System.Configuration.ConfigurationManager.AppSettings["tor_email_ar_overdue_cc"], sSalesPerson[0], arrayofSalesReps);
                                }
                                if (Program.Move.ToUpper() == "Y")
                                {
                                    if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                                        Utilities.MoveFile(fullFileName, Program.torbackupPath, System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"]);

                                    if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                                        Utilities.MoveFile(fullFileName, Program.vanbackupPath, System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data_cc"]);
                                }
                                return "OK";
                            }
                            else
                            {
                                saveBodyText = bodyText;
                                if ((status == "ONHOLD") || (status == "ONHOLD_AND_OVERDUE"))
                                {
                                    sSubject = "ON HOLD ACCT - " + custNo + " - " + custName + " - " + newOrderNumber + " - " + newShipmentNumber;
                                    bodyText = bodyText + Environment.NewLine + "Customer on hold." + Environment.NewLine;
                                    bodyText = bodyText + Environment.NewLine + "File - " + fullFileName + Environment.NewLine;
                                    bodyText = bodyText + Environment.NewLine + "Customer - " + custNo + Environment.NewLine;
                                    bodyText = bodyText + Environment.NewLine + "Order Number - " + newOrderNumber + Environment.NewLine;
                                    bodyText = bodyText + Environment.NewLine + "Shipment Number - " + newShipmentNumber + Environment.NewLine;
                                    bodyText = bodyText + Environment.NewLine + "Reference - " + reference + Environment.NewLine;
                                    if (sSalesRepLoc == "MTL")
                                        email.Sendemail(sSubject, bodyText, null, System.Configuration.ConfigurationManager.AppSettings["mtl_email_on_hold"], System.Configuration.ConfigurationManager.AppSettings["tor_email_on_hold_cc"], sSalesPerson[0], arrayofSalesReps);
                                    else if (sSalesRepLoc == "TOR")
                                        email.Sendemail(sSubject, bodyText, null, System.Configuration.ConfigurationManager.AppSettings["tor_email_on_hold"], System.Configuration.ConfigurationManager.AppSettings["tor_email_on_hold_cc"], sSalesPerson[0], arrayofSalesReps);
                                    else if (sSalesRepLoc == "VAN")
                                        email.Sendemail(sSubject, bodyText, null, System.Configuration.ConfigurationManager.AppSettings["van_email_on_hold"], System.Configuration.ConfigurationManager.AppSettings["van_email_on_hold_cc"], sSalesPerson[0], arrayofSalesReps);
                                    else if (sSalesRepLoc == "CAL")
                                        email.Sendemail(sSubject, bodyText, null, System.Configuration.ConfigurationManager.AppSettings["cal_email_on_hold"], System.Configuration.ConfigurationManager.AppSettings["van_email_on_hold_cc"], sSalesPerson[0], arrayofSalesReps);
                                    else
                                        email.Sendemail(sSubject, bodyText, null, System.Configuration.ConfigurationManager.AppSettings["tor_email_on_hold"], System.Configuration.ConfigurationManager.AppSettings["tor_email_on_hold_cc"], sSalesPerson[0], arrayofSalesReps);
                                }

                                if ((status == "OVERDUE") || (status == "ONHOLD_AND_OVERDUE"))
                                {
                                    sSubject = "OVERDUE INVOICES - " + custNo + " - " + custName + " - " + newOrderNumber + " - " + newShipmentNumber;
                                    bodyText = saveBodyText;
                                    bodyText = bodyText + Environment.NewLine + "Customer has overdue invoices greater than " + Program.DaysOverDue.ToString() + " days." + Environment.NewLine;
                                    bodyText = bodyText + Environment.NewLine + "File - " + fullFileName + Environment.NewLine;
                                    bodyText = bodyText + Environment.NewLine + "Customer - " + custNo + Environment.NewLine;
                                    bodyText = bodyText + Environment.NewLine + "Order Number - " + newOrderNumber + Environment.NewLine;
                                    bodyText = bodyText + Environment.NewLine + "Shipment Number - " + newShipmentNumber + Environment.NewLine;
                                    bodyText = bodyText + Environment.NewLine + "Reference - " + reference + Environment.NewLine;
                                    if (sSalesRepLoc == "MTL")
                                        email.Sendemail(sSubject, bodyText, null, System.Configuration.ConfigurationManager.AppSettings["mtl_email_ar_overdue"], System.Configuration.ConfigurationManager.AppSettings["tor_email_ar_overdue_cc"], sSalesPerson[0], arrayofSalesReps);
                                    else if (sSalesRepLoc == "TOR")
                                        email.Sendemail(sSubject, bodyText, null, System.Configuration.ConfigurationManager.AppSettings["tor_email_ar_overdue"], System.Configuration.ConfigurationManager.AppSettings["tor_email_ar_overdue_cc"], sSalesPerson[0], arrayofSalesReps);
                                    else if (sSalesRepLoc == "VAN")
                                        email.Sendemail(sSubject, bodyText, null, System.Configuration.ConfigurationManager.AppSettings["van_email_ar_overdue"], System.Configuration.ConfigurationManager.AppSettings["van_email_ar_overdue_cc"], sSalesPerson[0], arrayofSalesReps);
                                    else if (sSalesRepLoc == "CAL")
                                        email.Sendemail(sSubject, bodyText, null, System.Configuration.ConfigurationManager.AppSettings["cal_email_ar_overdue"], System.Configuration.ConfigurationManager.AppSettings["van_email_ar_overdue_cc"], sSalesPerson[0], arrayofSalesReps);
                                    else
                                        email.Sendemail(sSubject, bodyText, null, System.Configuration.ConfigurationManager.AppSettings["tor_email_ar_overdue"], System.Configuration.ConfigurationManager.AppSettings["tor_email_ar_overdue_cc"], sSalesPerson[0], arrayofSalesReps);
                                }

                                if (Program.Move.ToUpper() == "Y")
                                {
                                    if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                                        Utilities.MoveFile(fullFileName, Program.torbackupPath, System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"]);

                                    if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                                        Utilities.MoveFile(fullFileName, Program.vanbackupPath, System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data_cc"]);
                                }
                                return "OK";
                            }
                        }
                        else
                        {
                            if (Program.Move.ToUpper() == "Y")
                            {
                                if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                                    Utilities.MoveFile(fullFileName, Program.torbackupPath, System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"]);

                                if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                                    Utilities.MoveFile(fullFileName, Program.vanbackupPath, System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data_cc"]);
                            }
                            //log.LogWrite(Program.backupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + custNo + ",FAILED TO POST," + reference);
                            bodyText = bodyText + Environment.NewLine + "Failed to create shipment.";
                            bodyText = bodyText + Environment.NewLine + "File - " + fullFileName;
                            bodyText = bodyText + Environment.NewLine + "Customer - " + custNo;
                            bodyText = bodyText + Environment.NewLine + "Reference - " + reference;
                            bodyText = bodyText + Environment.NewLine + "Order - " + newOrderNumber;
                            email.Sendemail("FAILED TO POST SHIPMENT", bodyText, null, System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"], null, null);
                            return "BAD";
                        }
                    }
                    else
                    {
                        if (Program.Move.ToUpper() == "Y")
                        {
                            if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                                Utilities.MoveFile(fullFileName, Program.torbackupPath, System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"]);

                            if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                                Utilities.MoveFile(fullFileName, Program.vanbackupPath, System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data_cc"]);
                        }

                        //log.LogWrite(Program.backupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + custNo + ",FAILED TO POST," + reference);
                        bodyText = bodyText + Environment.NewLine + "Failed to Post.";
                        bodyText = bodyText + Environment.NewLine + "File - " + fullFileName;
                        bodyText = bodyText + Environment.NewLine + "Customer - " + custNo;
                        bodyText = bodyText + Environment.NewLine + "Reference - " + reference;
                        email.Sendemail("FAILED TO POST", bodyText, null, System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"], null, null);
                        return "BAD";
                    }
                }
                else
                {
                    if (Program.Move.ToUpper() == "Y")
                    {
                        if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                            Utilities.MoveFile(fullFileName, Program.torbackupPath, System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"]);

                        if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                            Utilities.MoveFile(fullFileName, Program.vanbackupPath, System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data_cc"]);
                    }
                    //log.LogWrite(Program.backupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + custNo + ",FAILED TO POST," + reference);
                    bodyText = bodyText + Environment.NewLine + "Failed to Post.";
                    bodyText = bodyText + Environment.NewLine + "File - " + fullFileName;
                    bodyText = bodyText + Environment.NewLine + "Customer - " + custNo;
                    bodyText = bodyText + Environment.NewLine + "Reference - " + reference;
                    email.Sendemail("FAILED TO POST", bodyText, null, System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"], null, null);
                    return "BAD";
                }
            }
            else
            {
                if (totalQtyShipped <= 0)
                {
                    if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                        log.LogWrite(Program.torbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + "No quantity on hand for all the items in location " + location + ".,File not imported");

                    if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                        log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + "No quantity on hand for all the items in location " + location + "., File not imported");
                }
                else
                {
                    if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                        log.LogWrite(Program.torbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + "File has bad data.,ERROR - File not imported");

                    if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                        log.LogWrite(Program.vanbackupPath, Program.txtOutputFile, Environment.UserName + "," + DateTime.Now + "," + fullFileName + "," + "File has bad data.,ERROR - File not imported");
                }

                if (Program.Move.ToUpper() == "Y")
                {
                    if ((Utilities.IsBitSet(Program.iLoc, 0)) || (Utilities.IsBitSet(Program.iLoc, 1)))    // Toronto or Montreal
                        Utilities.MoveFile(fullFileName, Program.torbackupPath, System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"]);

                    if ((Utilities.IsBitSet(Program.iLoc, 2)) || (Utilities.IsBitSet(Program.iLoc, 3)))    // Vancouver or Calgary
                        Utilities.MoveFile(fullFileName, Program.vanbackupPath, System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data_cc"]);
                }
                bodyText = bodyText + Environment.NewLine + "Issues with 1 or more items." + Environment.NewLine;
                bodyText = bodyText + Environment.NewLine + "File - " + fullFileName + Environment.NewLine;
                bodyText = bodyText + Environment.NewLine + "Customer - " + custNo + Environment.NewLine;
                bodyText = bodyText + Environment.NewLine + "Reference - " + reference + Environment.NewLine;
                email.Sendemail("FAILED TO POST - Issue with 1 or more items", bodyText, null, System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"], null, null);
                return "BAD";
            }
        }
    }
}

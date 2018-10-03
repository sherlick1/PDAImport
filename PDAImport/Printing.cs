using System;
using System.Globalization;
using System.IO;
using System.Linq;
using ACCPAC.Advantage;


namespace PDAImport
{
    class Printing
    {

        public static void PrintPickingSlip(string shipmentNumber, string sLoc, ACCPAC.Advantage.Session sageSession)
        {
//            bool temp;
//            string pick_filename;
            ACCPAC.Advantage.Report rpt;
            ACCPAC.Advantage.PrintSetup rptPrintSetup;


            if (Program.sourceSystem == "JFC")
            {
                if ((sLoc == "VAN") || (sLoc == "CAL") || (sLoc == "VANCAL"))
                    rpt = sageSession.ReportSelect(@"OEPICK01[\\192.168.190.20\Sage300\SharedData\Custom\OE63A\ENG\VAN_PICKSLIP_PDA.RPT]", "      ", "      ");
                else
                    rpt = sageSession.ReportSelect(@"OEPICK01[\\192.168.190.20\Sage300\SharedData\Custom\OE63A\ENG\TOR_PICKSLIP_PDA.RPT]", "      ", "      ");
            }
            else
                rpt = sageSession.ReportSelect(@"OEPICK01[C:\SAGE300\OE63A\ENG\TOR_PICKSLIP_LASER.RPT]", "      ", "      ");

            rptPrintSetup = sageSession.GetPrintSetup("      ", "      ");

            if (Program.sourceSystem == "JFC")
                if ((sLoc == "VAN") || (sLoc == "CAL") || (sLoc == "VANCAL"))
                {
                    rptPrintSetup.DeviceName = "Fujitsu DL3850+ (Vancouver)";
                    rptPrintSetup.OutputName = "192.168.191.161";
                    rptPrintSetup.PaperSize = 203;   // Dot Matrix Printer
                }
                else
                {
                    rptPrintSetup.DeviceName = "Fujitsu DL3850+";
                    rptPrintSetup.OutputName = "192.168.190.13";
                    rptPrintSetup.PaperSize = 203;   // Dot Matrix Printer
                }
            else
            {
                rptPrintSetup.DeviceName = "@\\dctor02\\TOR-HP3";
                rptPrintSetup.OutputName = "TOR-HP3";
                rptPrintSetup.PaperSize = 1;    // Laser Printer
            }

            rptPrintSetup.Orientation = 1;
            rptPrintSetup.PaperSource = 15;
            rpt.PrinterSetup(rptPrintSetup);

            rpt.SetParam("SELECTBY", "1");              // Report parameter: 2
            rpt.SetParam("SORTBY", "0");                // Report parameter: 3
            rpt.SetParam("FROMSELECT", shipmentNumber); // Report parameter: 4
            rpt.SetParam("TOSELECT", shipmentNumber);   // Report parameter: 5
            rpt.SetParam("FROMLOC", " ");               // Report parameter: 6
            rpt.SetParam("TOLOC", "ZZZZZZ");            // Report parameter: 7
            rpt.SetParam("PRINTBY", "0");               // Report parameter: 14
            rpt.SetParam("SERIALLOTNUMBERS", "0");      // Report parameter: 15
            rpt.SetParam("PRINTKIT", "0");              // Report parameter: 11
            rpt.SetParam("PRINTBOM", "0");              // Report parameter: 12
            rpt.SetParam("REPRINT", "0");               // Report parameter: 8
            rpt.SetParam("QTYDEC", "4");                // Report parameter: 9 - O/E Sales History:Detail,Sort by Item Number
            rpt.SetParam("COMPLETED", "0");             // Report parameter: 10

            rpt.NumberOfCopies = 1;

            // Print as PDF
            // rpt.Destination = PrintDestination.File;
            // rpt.Format = PrintFormat.PDF;
            // Pick_filename = @"d:\Temp\" & ShipmentNumber & ".pdf";
            // rpt.PrintDirectory = Pick_filename;
            // rpt.Print;

            if (Program.sMode != "Silent")
                if (Program.sourceSystem == "JFC")
                {
                    if (Program.sPrd.ToUpper() != "YES")
                        if (Program.output == "printer")
                            rpt.Destination = PrintDestination.Printer;
                        else
                            rpt.Destination = PrintDestination.Preview;

                    rpt.PrintDirectory = "";
                }
                else
                {
                    rpt.Destination = PrintDestination.Preview;
                    rpt.PrintDirectory = "";
                }

            else
            {
                rpt.Destination = PrintDestination.Printer;
                rpt.PrintDirectory = "";
            }

            rpt.Print();

            rpt.Dispose();
        }
    }
}

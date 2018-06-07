using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PDAImport
{
    public partial class Form1 : Form
    {
        private ACCPAC.Advantage.Session session;        
        private ACCPAC.Advantage.DBLink mDBLinkCmpRW;
        private ACCPAC.Advantage.View csView;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            short periods, qtr4Period, quarter;
            bool bActive, bRet;
            string year;
            DateTime startDate, endDate;

            session = new ACCPAC.Advantage.Session();
            session.Init("", "XX", "XX1000", "63A");
            session.Open("ADMIN", "RTFM", "JFJDAT", DateTime.Today, 0);
            mDBLinkCmpRW = session.OpenDBLink(ACCPAC.Advantage.DBLinkType.Company, ACCPAC.Advantage.DBLinkFlags.ReadWrite);

            ACCPAC.Advantage.Company compInfo = mDBLinkCmpRW.Company;
            textBox1.AppendText(compInfo.Name + " " + compInfo.HomeCurrency + " " + compInfo.LegalName + "\r\n");

            textBox1.AppendText("Sage 300 Version: "+ session.ACCPACVersion + " build version: " + session.ACCPACVersionBuild + " app: " + session.AppID + "app ver: " + session.AppVersion + "\r\n");
            textBox1.AppendText("user id: " + session.UserID + " language: " + session.UserLanguage + "\r\n");

            ACCPAC.Advantage.FiscalCalendar fiscCal = mDBLinkCmpRW.FiscalCalendar;

            bRet = fiscCal.GetYear("2019", out periods, out qtr4Period, out bActive);
            textBox1.AppendText("2019 periods = " + periods.ToString() + " qtrPeriod = " + qtr4Period.ToString() + "\r\n");
            bRet = fiscCal.GetYearDates("2019", out startDate, out endDate);
            textBox1.AppendText("2019 start/end date = " + startDate.ToString() + " " + endDate.ToString() + "\r\n");

            bRet = fiscCal.GetFirstYear(out year, out periods, out qtr4Period, out bActive);
            textBox1.AppendText("First year = " + year + " periods: " + periods.ToString() + " qtrPeriod = " + qtr4Period.ToString() + "\r\n");
            bRet = fiscCal.GetLastYear(out year, out periods, out qtr4Period, out bActive);
            textBox1.AppendText("Last year = " + year + " periods: " + periods.ToString() + " qtrPeriod = " + qtr4Period.ToString() + "\r\n");

            bRet = fiscCal.GetPeriod(new DateTime(2019, 5, 23), out periods, out year, out bActive);
            textBox1.AppendText("Fisc infor for 2019/05/23 = " + year + " period: " + periods.ToString() + " open: " + bActive.ToString() + "\r\n");
            bRet = fiscCal.GetPeriodDates("2019", 5, out startDate, out endDate, out bActive);
            textBox1.AppendText("Period Dates for 2019/5 = " + startDate.ToString() + " to: " + endDate.ToString() + " open: " + bActive.ToString() + "\r\n");

            bRet = fiscCal.GetQuarter("2019", 5, out quarter);
            textBox1.AppendText("Quarter for 2019/5 = " + quarter.ToString() + "\r\n");
            bRet = fiscCal.GetQuarterDates("2019", 3, out startDate, out endDate);
            textBox1.AppendText("2019 Quarter 3 start/end dates: " + startDate.ToString() + " " + endDate.ToString() + "\r\n");

            bRet = fiscCal.DatesFromPeriod(5, ACCPAC.Advantage.FiscalPeriodType.Monthly, 1, new DateTime(2019, 1, 1), out startDate, out endDate);
            textBox1.AppendText("Period dates for 2019/5: " + startDate.ToString() + " " + endDate.ToString() + "\r\n");
            bRet = fiscCal.DateToPeriod(new DateTime(2019, 5, 1), ACCPAC.Advantage.FiscalPeriodType.Monthly, 1, new DateTime(2019, 1, 1), out periods);
            textBox1.AppendText("Period for 2019/5/1: " + periods.ToString() + "\r\n");

            ACCPAC.Advantage.Currency curInfo = mDBLinkCmpRW.GetCurrency("CAD");

            textBox1.AppendText("DBLink currency: " + curInfo.Code + " " + curInfo.Description + " " + curInfo.Symbol + "\r\n");
            textBox1.AppendText(curInfo.Decimals.ToString() + " " + curInfo.ThousandSeparator.ToString() + " " + curInfo.NegativeDisplay.ToString() + "\r\n");

            textBox1.AppendText("Num active apps: " + mDBLinkCmpRW.ActiveApplications.Count.ToString()+"\r\n");
            ACCPAC.Advantage.ActiveApplication appInfo = mDBLinkCmpRW.ActiveApplications[3];
            textBox1.AppendText("app id: " + appInfo.AppID + " version: " + appInfo.AppVersion+ " sequence: " + appInfo.Sequence + "\r\n");
            textBox1.AppendText("data leve: " + appInfo.DataLevel.ToString() + " installed: " + appInfo.IsInstalled.ToString() + "\r\n");

            ACCPAC.Advantage.CurrencyTable curTab = mDBLinkCmpRW.GetCurrencyTable("USD", "SP");
            textBox1.AppendText("Currency Table: " +curTab.CurrencyCode + " " +curTab.RateType + " Desc: " +curTab.Description +" Source: "+curTab.SourceOfRates+ "\r\n");

            string rateTypeDesc;
            mDBLinkCmpRW.GetCurrencyRateTypeDescription("SP", out rateTypeDesc);
            textBox1.AppendText("SP Rate Type Description: " + rateTypeDesc + "\r\n");

            System.DateTime rateDate = new DateTime(2017, 02, 27);       
            ACCPAC.Advantage.CurrencyRate curRate = mDBLinkCmpRW.GetCurrencyRate("CAD", "SP", "USD", rateDate);
            textBox1.AppendText(curRate.HomeCurrency + " " + curRate.RateType + " " + curRate.SourceCurrency + " " +curRate.Rate + "\r\n" );

            csView = mDBLinkCmpRW.OpenView("CS0001");
            csView.Fetch(false);
            string tmpStr = csView.Fields.FieldByID(1).Value + "  " + csView.Fields.FieldByID(2).Value;
            label1.Text = tmpStr;
            textBox1.AppendText(tmpStr + "\r\n");


        }
    }
}

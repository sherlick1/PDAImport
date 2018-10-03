using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ACCPAC.Advantage;
using System.IO;  // Used only for the FileSystemWatcher Example

namespace PDAImport
{
    public partial class Form1 : Form
    {

        /**********************************************
         * Declarations
         **********************************************/

        public Form1()
        {
            InitializeComponent();
        }

        private void startButton_Click(object sender, EventArgs e)
        {
            CreateOrder createOrder = new CreateOrder();
            WireEventHandlers(createOrder);

            string message;
            string caption;
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            DialogResult result;

            if (TorButton.Checked)
            {
                Program.sLoc = "TOR";
                Program.torbackupPath = System.Configuration.ConfigurationManager.AppSettings["tor_backup_path"];
                Program.iLoc = Program.iLoc ^ 1;
            }
            else if (TorandMtlButton.Checked)
            {
                Program.sLoc = "TORMTL";
                Program.torbackupPath = System.Configuration.ConfigurationManager.AppSettings["tor_backup_path"];
                Program.iLoc = Program.iLoc ^ 1;
                Program.iLoc = Program.iLoc ^ 2;

            }
            else if (MtlButton.Checked)
            {
                Program.sLoc = "MTL";
                Program.torbackupPath = System.Configuration.ConfigurationManager.AppSettings["tor_backup_path"];
                Program.iLoc = Program.iLoc ^ 2;
            }
            else if (VanButton.Checked)
            {
                Program.sLoc = "VAN";
                Program.vanbackupPath = System.Configuration.ConfigurationManager.AppSettings["van_backup_path"];
                Program.iLoc = Program.iLoc ^ 4;
            }
            else if (CalButton.Checked)
            {
                Program.sLoc = "CAL";
                Program.vanbackupPath = System.Configuration.ConfigurationManager.AppSettings["van_backup_path"];
                Program.iLoc = Program.iLoc ^ 8;
            }
            else if (VanandCalButton.Checked)
            {
                Program.sLoc = "VANCAL";
                Program.vanbackupPath = System.Configuration.ConfigurationManager.AppSettings["van_backup_path"];
                Program.iLoc = Program.iLoc ^ 4;
                Program.iLoc = Program.iLoc ^ 8;
            }

            if (printToPrinter.Checked)
                Program.output = "printer";
            else
                Program.output = "preview";

            if (emailSalesRep.Checked)
                Program.emailSalesRep = true;
            else
                Program.emailSalesRep = false;

            if (createOrder.PerformPDAImport())
            {
                message = "Import Finished Successfully.";
            }
            else
            {
                message = "Import had errors!";
            }

            if (TorButton.Checked)
            {
                Utilities.CopyFile(Program.torbackupPath, Program.txtOutputFile, System.Configuration.ConfigurationManager.AppSettings["tor_error_path"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"]);
            }
            else if (TorandMtlButton.Checked)
            {
                Utilities.CopyFile(Program.torbackupPath, Program.txtOutputFile, System.Configuration.ConfigurationManager.AppSettings["tor_error_path"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"]);
            }
            else if (MtlButton.Checked)
            {
                Utilities.CopyFile(Program.torbackupPath, Program.txtOutputFile, System.Configuration.ConfigurationManager.AppSettings["tor_error_path"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["tor_email_bad_data_cc"]);
            }
            else if (VanButton.Checked)
            {
                Utilities.CopyFile(Program.vanbackupPath, Program.txtOutputFile, System.Configuration.ConfigurationManager.AppSettings["van_error_path"], System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data_cc"]);
            }
            else if (CalButton.Checked)
            {
                Utilities.CopyFile(Program.vanbackupPath, Program.txtOutputFile, System.Configuration.ConfigurationManager.AppSettings["van_error_path"], System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data_cc"]);
            }
            else if (VanandCalButton.Checked)
            {
                Utilities.CopyFile(Program.vanbackupPath, Program.txtOutputFile, System.Configuration.ConfigurationManager.AppSettings["van_error_path"], System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data"], System.Configuration.ConfigurationManager.AppSettings["van_email_bad_data_cc"]);
            }

            caption = "Status of Import";
            result = MessageBox.Show(message, caption, buttons);

            this.Close();

        }

        private void WireEventHandlers(CreateOrder createOrder)
        {
            MyHandler1 handler = new MyHandler1(OnHandler1);
            createOrder.Event1 += handler;
        }

        public void OnHandler1(object sender, MyEvent e)
        {
            this.Results.Text = e.message;
            this.StatusLabel.Text = e.message;
            this.statusStrip1.Update();
        }

        private void Closebutton_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            switch (Program.sLoc)
            {
                case "TOR":
                    TorButton.Checked = true;
                    break;
                case "MTL":
                    MtlButton.Checked = true;
                    break;
                case "TORMTL":
                    TorandMtlButton.Checked = true;
                    break;
                case "VAN":
                    VanButton.Checked = true;
                    break;
                case "CAL":
                    CalButton.Checked = true;
                    break;
                case "VANCAL":
                    VanandCalButton.Checked = true;
                    break;
            }

            if (Program.sPrd.ToUpper() == "PRD")
            {
                printToPrinter.Checked = true;
                emailSalesRep.Checked = true;
            }

            StatusLabel.Text = String.Empty;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}

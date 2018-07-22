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

        private void button1_Click(object sender, EventArgs e)
        {
            CreateOrder createOrder = new CreateOrder();
            WireEventHandlers(createOrder);

            string message;
            string caption;
            string sLoc = string.Empty; 
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            DialogResult result;

            if (TorButton.Checked)
                sLoc = "TOR";
            else if (TorandMtlButton.Checked)
                sLoc = "TORMTL";
            else if (MtlButton.Checked)
                sLoc = "MTL";
            else if (VanButton.Checked)
                sLoc = "VAN";
            else if (CalButton.Checked)
                sLoc = "CAL";
            else if (VanandCalButton.Checked)
                sLoc = "VANCAL";

            if (createOrder.PerformPDAImport(sLoc))
            {
                message = "Import Finished Successfully.";
            }
            else
            {
                message = "Import had errors!";
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
                    TorandMtlButton.Checked = true;
                    break;
                case "VAN":
                    VanButton.Checked = true;
                    break;
            }

            StatusLabel.Text = String.Empty;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}

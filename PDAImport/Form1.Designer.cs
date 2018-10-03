namespace PDAImport
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.StartButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.TorButton = new System.Windows.Forms.RadioButton();
            this.MtlButton = new System.Windows.Forms.RadioButton();
            this.TorandMtlButton = new System.Windows.Forms.RadioButton();
            this.VanButton = new System.Windows.Forms.RadioButton();
            this.button2 = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.StatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.Results = new System.Windows.Forms.Label();
            this.CalButton = new System.Windows.Forms.RadioButton();
            this.VanandCalButton = new System.Windows.Forms.RadioButton();
            this.printToPrinter = new System.Windows.Forms.CheckBox();
            this.emailSalesRep = new System.Windows.Forms.CheckBox();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // StartButton
            // 
            this.StartButton.Location = new System.Drawing.Point(39, 224);
            this.StartButton.Name = "StartButton";
            this.StartButton.Size = new System.Drawing.Size(75, 23);
            this.StartButton.TabIndex = 0;
            this.StartButton.Text = "Start";
            this.StartButton.UseVisualStyleBackColor = true;
            this.StartButton.Click += new System.EventHandler(this.startButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(39, 34);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(109, 14);
            this.label1.TabIndex = 1;
            this.label1.Text = "Locations to Process";
            // 
            // TorButton
            // 
            this.TorButton.AutoSize = true;
            this.TorButton.Location = new System.Drawing.Point(42, 66);
            this.TorButton.Name = "TorButton";
            this.TorButton.Size = new System.Drawing.Size(114, 17);
            this.TorButton.TabIndex = 2;
            this.TorButton.TabStop = true;
            this.TorButton.Text = "Toronto Shipments";
            this.TorButton.UseVisualStyleBackColor = true;
            // 
            // MtlButton
            // 
            this.MtlButton.AutoSize = true;
            this.MtlButton.Location = new System.Drawing.Point(42, 89);
            this.MtlButton.Name = "MtlButton";
            this.MtlButton.Size = new System.Drawing.Size(118, 17);
            this.MtlButton.TabIndex = 3;
            this.MtlButton.TabStop = true;
            this.MtlButton.Text = "Montreal Shipments";
            this.MtlButton.UseVisualStyleBackColor = true;
            // 
            // TorandMtlButton
            // 
            this.TorandMtlButton.AutoSize = true;
            this.TorandMtlButton.Location = new System.Drawing.Point(42, 112);
            this.TorandMtlButton.Name = "TorandMtlButton";
            this.TorandMtlButton.Size = new System.Drawing.Size(179, 17);
            this.TorandMtlButton.TabIndex = 4;
            this.TorandMtlButton.TabStop = true;
            this.TorandMtlButton.Text = "Toronto and Montreal Shipments";
            this.TorandMtlButton.UseVisualStyleBackColor = true;
            // 
            // VanButton
            // 
            this.VanButton.AutoSize = true;
            this.VanButton.Location = new System.Drawing.Point(42, 135);
            this.VanButton.Name = "VanButton";
            this.VanButton.Size = new System.Drawing.Size(129, 17);
            this.VanButton.TabIndex = 5;
            this.VanButton.TabStop = true;
            this.VanButton.Text = "Vancouver Shipments";
            this.VanButton.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(171, 224);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 6;
            this.button2.Text = "Close";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.Closebutton_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.StatusLabel});
            this.statusStrip1.Location = new System.Drawing.Point(0, 338);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(565, 22);
            this.statusStrip1.TabIndex = 8;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // StatusLabel
            // 
            this.StatusLabel.Name = "StatusLabel";
            this.StatusLabel.Size = new System.Drawing.Size(118, 17);
            this.StatusLabel.Text = "toolStripStatusLabel1";
            // 
            // Results
            // 
            this.Results.BackColor = System.Drawing.Color.White;
            this.Results.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.Results.Location = new System.Drawing.Point(271, 57);
            this.Results.Name = "Results";
            this.Results.Size = new System.Drawing.Size(242, 61);
            this.Results.TabIndex = 9;
            // 
            // CalButton
            // 
            this.CalButton.AutoSize = true;
            this.CalButton.Location = new System.Drawing.Point(42, 158);
            this.CalButton.Name = "CalButton";
            this.CalButton.Size = new System.Drawing.Size(112, 17);
            this.CalButton.TabIndex = 10;
            this.CalButton.TabStop = true;
            this.CalButton.Text = "Calgary Shipments";
            this.CalButton.UseVisualStyleBackColor = true;
            this.CalButton.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // VanandCalButton
            // 
            this.VanandCalButton.AutoSize = true;
            this.VanandCalButton.Location = new System.Drawing.Point(42, 181);
            this.VanandCalButton.Name = "VanandCalButton";
            this.VanandCalButton.Size = new System.Drawing.Size(188, 17);
            this.VanandCalButton.TabIndex = 11;
            this.VanandCalButton.TabStop = true;
            this.VanandCalButton.Text = "Vancouver and Calgary Shipments";
            this.VanandCalButton.UseVisualStyleBackColor = true;
            // 
            // printToPrinter
            // 
            this.printToPrinter.AutoSize = true;
            this.printToPrinter.Location = new System.Drawing.Point(271, 158);
            this.printToPrinter.Name = "printToPrinter";
            this.printToPrinter.Size = new System.Drawing.Size(92, 17);
            this.printToPrinter.TabIndex = 12;
            this.printToPrinter.Text = "Print to Printer";
            this.printToPrinter.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.printToPrinter.UseVisualStyleBackColor = true;
            // 
            // emailSalesRep
            // 
            this.emailSalesRep.AutoSize = true;
            this.emailSalesRep.Location = new System.Drawing.Point(271, 181);
            this.emailSalesRep.Name = "emailSalesRep";
            this.emailSalesRep.Size = new System.Drawing.Size(102, 17);
            this.emailSalesRep.TabIndex = 13;
            this.emailSalesRep.Text = "email Sales Rep";
            this.emailSalesRep.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.emailSalesRep.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(565, 360);
            this.Controls.Add(this.emailSalesRep);
            this.Controls.Add(this.printToPrinter);
            this.Controls.Add(this.VanandCalButton);
            this.Controls.Add(this.CalButton);
            this.Controls.Add(this.Results);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.VanButton);
            this.Controls.Add(this.TorandMtlButton);
            this.Controls.Add(this.MtlButton);
            this.Controls.Add(this.TorButton);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.StartButton);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button StartButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton TorButton;
        private System.Windows.Forms.RadioButton MtlButton;
        private System.Windows.Forms.RadioButton TorandMtlButton;
        private System.Windows.Forms.RadioButton VanButton;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel StatusLabel;
        private System.Windows.Forms.Label Results;
        private System.Windows.Forms.RadioButton CalButton;
        private System.Windows.Forms.RadioButton VanandCalButton;
        private System.Windows.Forms.CheckBox printToPrinter;
        private System.Windows.Forms.CheckBox emailSalesRep;
    }

}


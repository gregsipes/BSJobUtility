namespace LawsonArchive
{
    partial class FrmMain
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
            this.TxtEmployeeData = new System.Windows.Forms.TextBox();
            this.CmbEmployee = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.RadEmployeID = new System.Windows.Forms.RadioButton();
            this.RadEmployeeName = new System.Windows.Forms.RadioButton();
            this.label2 = new System.Windows.Forms.Label();
            this.MnuMain = new System.Windows.Forms.MenuStrip();
            this.MnuFile = new System.Windows.Forms.ToolStripMenuItem();
            this.MnuPrintPreview = new System.Windows.Forms.ToolStripMenuItem();
            this.MnuPrint = new System.Windows.Forms.ToolStripMenuItem();
            this.MnuExit = new System.Windows.Forms.ToolStripMenuItem();
            this.label3 = new System.Windows.Forms.Label();
            this.CmbYear = new System.Windows.Forms.ComboBox();
            this.TxtWages = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.MnuMain.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // TxtEmployeeData
            // 
            this.TxtEmployeeData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TxtEmployeeData.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtEmployeeData.Location = new System.Drawing.Point(6, 53);
            this.TxtEmployeeData.Multiline = true;
            this.TxtEmployeeData.Name = "TxtEmployeeData";
            this.TxtEmployeeData.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.TxtEmployeeData.Size = new System.Drawing.Size(723, 600);
            this.TxtEmployeeData.TabIndex = 0;
            this.TxtEmployeeData.WordWrap = false;
            // 
            // CmbEmployee
            // 
            this.CmbEmployee.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.CmbEmployee.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.CmbEmployee.FormattingEnabled = true;
            this.CmbEmployee.Location = new System.Drawing.Point(72, 27);
            this.CmbEmployee.Name = "CmbEmployee";
            this.CmbEmployee.Size = new System.Drawing.Size(405, 21);
            this.CmbEmployee.TabIndex = 1;
            this.CmbEmployee.SelectedIndexChanged += new System.EventHandler(this.CmbEmployee_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Employee";
            // 
            // RadEmployeID
            // 
            this.RadEmployeID.AutoSize = true;
            this.RadEmployeID.Location = new System.Drawing.Point(58, 3);
            this.RadEmployeID.Name = "RadEmployeID";
            this.RadEmployeID.Size = new System.Drawing.Size(81, 17);
            this.RadEmployeID.TabIndex = 3;
            this.RadEmployeID.Text = "Employee #";
            this.RadEmployeID.UseVisualStyleBackColor = true;
            this.RadEmployeID.CheckedChanged += new System.EventHandler(this.RadEmployeID_CheckedChanged);
            // 
            // RadEmployeeName
            // 
            this.RadEmployeeName.AutoSize = true;
            this.RadEmployeeName.Checked = true;
            this.RadEmployeeName.Location = new System.Drawing.Point(145, 3);
            this.RadEmployeeName.Name = "RadEmployeeName";
            this.RadEmployeeName.Size = new System.Drawing.Size(102, 17);
            this.RadEmployeeName.TabIndex = 4;
            this.RadEmployeeName.TabStop = true;
            this.RadEmployeeName.Text = "Employee Name";
            this.RadEmployeeName.UseVisualStyleBackColor = true;
            this.RadEmployeeName.CheckedChanged += new System.EventHandler(this.RadEmployeeName_CheckedChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(0, 5);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Search By";
            // 
            // MnuMain
            // 
            this.MnuMain.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.MnuMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.MnuFile,
            this.MnuExit});
            this.MnuMain.Location = new System.Drawing.Point(0, 0);
            this.MnuMain.Name = "MnuMain";
            this.MnuMain.Padding = new System.Windows.Forms.Padding(4, 2, 0, 2);
            this.MnuMain.Size = new System.Drawing.Size(1081, 24);
            this.MnuMain.TabIndex = 6;
            this.MnuMain.Text = "MainMenu";
            // 
            // MnuFile
            // 
            this.MnuFile.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.MnuPrintPreview,
            this.MnuPrint});
            this.MnuFile.Name = "MnuFile";
            this.MnuFile.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.F)));
            this.MnuFile.Size = new System.Drawing.Size(37, 20);
            this.MnuFile.Text = "&File";
            // 
            // MnuPrintPreview
            // 
            this.MnuPrintPreview.Name = "MnuPrintPreview";
            this.MnuPrintPreview.Size = new System.Drawing.Size(170, 22);
            this.MnuPrintPreview.Text = "Print Pre&view";
            this.MnuPrintPreview.Click += new System.EventHandler(this.MnuPrintPreview_Click);
            // 
            // MnuPrint
            // 
            this.MnuPrint.Name = "MnuPrint";
            this.MnuPrint.Size = new System.Drawing.Size(170, 22);
            this.MnuPrint.Text = "&Print";
            this.MnuPrint.Click += new System.EventHandler(this.MnuPrint_Click);
            // 
            // MnuExit
            // 
            this.MnuExit.Name = "MnuExit";
            this.MnuExit.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.X)));
            this.MnuExit.Size = new System.Drawing.Size(38, 20);
            this.MnuExit.Text = "E&xit";
            this.MnuExit.Click += new System.EventHandler(this.MnuExit_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(735, 30);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(63, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Payroll Year";
            // 
            // CmbYear
            // 
            this.CmbYear.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.CmbYear.FormattingEnabled = true;
            this.CmbYear.Items.AddRange(new object[] {
            "2008",
            "2007",
            "2006",
            "2005",
            "2004",
            "2003",
            "2002",
            "2001",
            "2000",
            "1999"});
            this.CmbYear.Location = new System.Drawing.Point(801, 27);
            this.CmbYear.MaxDropDownItems = 12;
            this.CmbYear.Name = "CmbYear";
            this.CmbYear.Size = new System.Drawing.Size(56, 21);
            this.CmbYear.TabIndex = 7;
            this.CmbYear.Text = "2008";
            this.CmbYear.SelectedIndexChanged += new System.EventHandler(this.CmbYear_SelectedIndexChanged);
            // 
            // TxtWages
            // 
            this.TxtWages.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TxtWages.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtWages.Location = new System.Drawing.Point(734, 53);
            this.TxtWages.Multiline = true;
            this.TxtWages.Name = "TxtWages";
            this.TxtWages.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.TxtWages.Size = new System.Drawing.Size(337, 600);
            this.TxtWages.TabIndex = 9;
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.RadEmployeeName);
            this.panel1.Controls.Add(this.RadEmployeID);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Location = new System.Drawing.Point(483, 25);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(246, 25);
            this.panel1.TabIndex = 10;
            // 
            // FrmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1081, 681);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.TxtWages);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.CmbYear);
            this.Controls.Add(this.MnuMain);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.CmbEmployee);
            this.Controls.Add(this.TxtEmployeeData);
            this.Name = "FrmMain";
            this.Text = "Lawson Archive";
            this.MnuMain.ResumeLayout(false);
            this.MnuMain.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox TxtEmployeeData;
        private System.Windows.Forms.ComboBox CmbEmployee;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton RadEmployeID;
        private System.Windows.Forms.RadioButton RadEmployeeName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.MenuStrip MnuMain;
        private System.Windows.Forms.ToolStripMenuItem MnuFile;
        private System.Windows.Forms.ToolStripMenuItem MnuPrintPreview;
        private System.Windows.Forms.ToolStripMenuItem MnuPrint;
        private System.Windows.Forms.ToolStripMenuItem MnuExit;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox CmbYear;
        private System.Windows.Forms.TextBox TxtWages;
        private System.Windows.Forms.Panel panel1;
    }
}


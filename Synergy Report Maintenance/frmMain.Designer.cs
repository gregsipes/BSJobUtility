namespace Synergy_Report_Maintenance
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
            this.CmbSynergyReportType = new System.Windows.Forms.ComboBox();
            this.LblSynergyReportType = new System.Windows.Forms.Label();
            this.CmbReportToRefresh = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.CmdRefresh = new System.Windows.Forms.Button();
            this.CmdExit = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // CmbSynergyReportType
            // 
            this.CmbSynergyReportType.FormattingEnabled = true;
            this.CmbSynergyReportType.Items.AddRange(new object[] {
            "AGEDET:         Account Lockbox - Age Analsyis Detail",
            "AGESUMM:     Account Lockbox - Age Analysis Summary (Sorted by District and Acct)" +
                "",
            "AGESUMM2:   Account Lockbox - Age Analysis Summary (Sorted by Acct Only)",
            "DLYDRAW:      Barron\'s Daily Draw Worksheet",
            "DEPOSITS:      Direct Deposits - Home Delivery Carrier",
            "DEPOSIT:        Direct Deposits - Home Delivery Carrier (w/Adjustments)",
            "PREAUTH:       Direct Deposits - Home Delivery Carrier Pre-notification",
            "GRACEOWE:   Grace Due",
            "GRACEWO:     Grace Writeoff",
            "UR:                  Unearned Revenue"});
            this.CmbSynergyReportType.Location = new System.Drawing.Point(12, 131);
            this.CmbSynergyReportType.MaxDropDownItems = 50;
            this.CmbSynergyReportType.Name = "CmbSynergyReportType";
            this.CmbSynergyReportType.Size = new System.Drawing.Size(675, 24);
            this.CmbSynergyReportType.TabIndex = 0;
            this.CmbSynergyReportType.SelectedIndexChanged += new System.EventHandler(this.CmbSynergyReportType_SelectedIndexChanged);
            // 
            // LblSynergyReportType
            // 
            this.LblSynergyReportType.AutoSize = true;
            this.LblSynergyReportType.Location = new System.Drawing.Point(16, 111);
            this.LblSynergyReportType.Name = "LblSynergyReportType";
            this.LblSynergyReportType.Size = new System.Drawing.Size(186, 17);
            this.LblSynergyReportType.TabIndex = 1;
            this.LblSynergyReportType.Text = "Select Synergy Report Type";
            // 
            // CmbReportToRefresh
            // 
            this.CmbReportToRefresh.FormattingEnabled = true;
            this.CmbReportToRefresh.Location = new System.Drawing.Point(12, 217);
            this.CmbReportToRefresh.MaxDropDownItems = 50;
            this.CmbReportToRefresh.Name = "CmbReportToRefresh";
            this.CmbReportToRefresh.Size = new System.Drawing.Size(675, 24);
            this.CmbReportToRefresh.TabIndex = 2;
            this.CmbReportToRefresh.SelectedIndexChanged += new System.EventHandler(this.CmbReportToRefresh_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 197);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(169, 17);
            this.label1.TabIndex = 3;
            this.label1.Text = "Select Report To Refresh";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(243, 33);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(206, 32);
            this.label2.TabIndex = 4;
            this.label2.Text = "Refresh Report";
            // 
            // CmdRefresh
            // 
            this.CmdRefresh.Location = new System.Drawing.Point(128, 277);
            this.CmdRefresh.Name = "CmdRefresh";
            this.CmdRefresh.Size = new System.Drawing.Size(148, 40);
            this.CmdRefresh.TabIndex = 5;
            this.CmdRefresh.Text = "Refresh Report";
            this.CmdRefresh.UseVisualStyleBackColor = true;
            this.CmdRefresh.Click += new System.EventHandler(this.CmdRefresh_Click);
            // 
            // CmdExit
            // 
            this.CmdExit.Location = new System.Drawing.Point(429, 277);
            this.CmdExit.Name = "CmdExit";
            this.CmdExit.Size = new System.Drawing.Size(148, 40);
            this.CmdExit.TabIndex = 6;
            this.CmdExit.Text = "Exit";
            this.CmdExit.UseVisualStyleBackColor = true;
            this.CmdExit.Click += new System.EventHandler(this.CmdExit_Click);
            // 
            // FrmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(741, 358);
            this.Controls.Add(this.CmdExit);
            this.Controls.Add(this.CmdRefresh);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.CmbReportToRefresh);
            this.Controls.Add(this.LblSynergyReportType);
            this.Controls.Add(this.CmbSynergyReportType);
            this.Name = "FrmMain";
            this.Text = "Synergy Report Maintenance / UR Report Maintenance";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmMain_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox CmbSynergyReportType;
        private System.Windows.Forms.Label LblSynergyReportType;
        private System.Windows.Forms.ComboBox CmbReportToRefresh;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button CmdRefresh;
        private System.Windows.Forms.Button CmdExit;
    }
}


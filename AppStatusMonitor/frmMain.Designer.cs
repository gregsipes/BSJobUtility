namespace AppStatusMonitor
{
    partial class frmMain
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
            this.components = new System.ComponentModel.Container();
            this.timUpdateStatus = new System.Windows.Forms.Timer(this.components);
            this.pnlMonitors = new System.Windows.Forms.Panel();
            this.SuspendLayout();
            // 
            // timUpdateStatus
            // 
            this.timUpdateStatus.Interval = 1;
            this.timUpdateStatus.Tick += new System.EventHandler(this.TimUpdateStatus_Tick);
            // 
            // pnlMonitors
            // 
            this.pnlMonitors.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlMonitors.AutoScroll = true;
            this.pnlMonitors.BackColor = System.Drawing.SystemColors.Desktop;
            this.pnlMonitors.Location = new System.Drawing.Point(0, -1);
            this.pnlMonitors.Margin = new System.Windows.Forms.Padding(4);
            this.pnlMonitors.Name = "pnlMonitors";
            this.pnlMonitors.Size = new System.Drawing.Size(932, 495);
            this.pnlMonitors.TabIndex = 1;
            this.pnlMonitors.Click += new System.EventHandler(this.PnlMonitors_Click);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(929, 491);
            this.Controls.Add(this.pnlMonitors);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "frmMain";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.Text = "Application Status Monitor";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmMain_FormClosing);
            this.ResizeEnd += new System.EventHandler(this.FrmMain_ResizeEnd);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Timer timUpdateStatus;
        private System.Windows.Forms.Panel pnlMonitors;
    }
}


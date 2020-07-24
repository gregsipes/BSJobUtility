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
            this.udAppMonitor = new AppStatusControl.AppStatusUserControl();
            this.pnlMonitors.SuspendLayout();
            this.SuspendLayout();
            // 
            // timUpdateStatus
            // 
            this.timUpdateStatus.Interval = 1;
            this.timUpdateStatus.Tick += new System.EventHandler(this.timUpdateStatus_Tick);
            // 
            // pnlMonitors
            // 
            this.pnlMonitors.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlMonitors.AutoScroll = true;
            this.pnlMonitors.BackColor = System.Drawing.SystemColors.Desktop;
            this.pnlMonitors.Controls.Add(this.udAppMonitor);
            this.pnlMonitors.Location = new System.Drawing.Point(0, -1);
            this.pnlMonitors.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.pnlMonitors.Name = "pnlMonitors";
            this.pnlMonitors.Size = new System.Drawing.Size(932, 495);
            this.pnlMonitors.TabIndex = 1;
            // 
            // udAppMonitor
            // 
            this.udAppMonitor.AppName = "<AppName>";
            this.udAppMonitor.BackColor = System.Drawing.SystemColors.Desktop;
            this.udAppMonitor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.udAppMonitor.Location = new System.Drawing.Point(51, 42);
            this.udAppMonitor.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.udAppMonitor.MaximumSize = new System.Drawing.Size(298, 64);
            this.udAppMonitor.MinimumSize = new System.Drawing.Size(298, 64);
            this.udAppMonitor.Name = "udAppMonitor";
            this.udAppMonitor.Size = new System.Drawing.Size(298, 64);
            this.udAppMonitor.TabIndex = 0;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(929, 491);
            this.Controls.Add(this.pnlMonitors);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "frmMain";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.Text = "Application Status Monitor";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmMain_FormClosing);
            this.ResizeEnd += new System.EventHandler(this.frmMain_ResizeEnd);
            this.pnlMonitors.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Timer timUpdateStatus;
        private System.Windows.Forms.Panel pnlMonitors;
        private AppStatusControl.AppStatusUserControl udAppMonitor;
    }
}


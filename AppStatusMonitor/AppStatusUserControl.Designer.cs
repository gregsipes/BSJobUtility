namespace AppStatusControl
{
    partial class AppStatusUserControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.LBLLastRunTime = new System.Windows.Forms.Label();
            this.LblAppName = new System.Windows.Forms.Label();
            this.LEDActivity = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.LEDActivity)).BeginInit();
            this.SuspendLayout();
            // 
            // LBLLastRunTime
            // 
            this.LBLLastRunTime.AutoSize = true;
            this.LBLLastRunTime.ForeColor = System.Drawing.Color.White;
            this.LBLLastRunTime.Location = new System.Drawing.Point(23, 44);
            this.LBLLastRunTime.Name = "LBLLastRunTime";
            this.LBLLastRunTime.Size = new System.Drawing.Size(116, 17);
            this.LBLLastRunTime.TabIndex = 16;
            this.LBLLastRunTime.Text = "<Last Run Time>";
            this.LBLLastRunTime.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // LblAppName
            // 
            this.LblAppName.AutoSize = true;
            this.LblAppName.ForeColor = System.Drawing.Color.White;
            this.LblAppName.Location = new System.Drawing.Point(23, 24);
            this.LblAppName.Name = "LblAppName";
            this.LblAppName.Size = new System.Drawing.Size(86, 17);
            this.LblAppName.TabIndex = 15;
            this.LblAppName.Text = "<AppName>";
            this.LblAppName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // LEDActivity
            // 
            this.LEDActivity.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.LEDActivity.Location = new System.Drawing.Point(4, 26);
            this.LEDActivity.Margin = new System.Windows.Forms.Padding(0);
            this.LEDActivity.Name = "LEDActivity";
            this.LEDActivity.Size = new System.Drawing.Size(12, 12);
            this.LEDActivity.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.LEDActivity.TabIndex = 9;
            this.LEDActivity.TabStop = false;
            // 
            // AppStatusUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Desktop;
            this.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Controls.Add(this.LBLLastRunTime);
            this.Controls.Add(this.LblAppName);
            this.Controls.Add(this.LEDActivity);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.MaximumSize = new System.Drawing.Size(298, 64);
            this.MinimumSize = new System.Drawing.Size(298, 64);
            this.Name = "AppStatusUserControl";
            this.Size = new System.Drawing.Size(298, 64);
            ((System.ComponentModel.ISupportInitialize)(this.LEDActivity)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label LBLLastRunTime;
        private System.Windows.Forms.Label LblAppName;
        private System.Windows.Forms.PictureBox LEDActivity;
    }
}

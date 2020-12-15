namespace SBSJournalEntryImport
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
            this.label1 = new System.Windows.Forms.Label();
            this.TxtImportFile = new System.Windows.Forms.TextBox();
            this.CmdBrowse = new System.Windows.Forms.Button();
            this.CmdExit = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.CmdImport = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(130, 87);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(120, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Select Excel Import File:";
            // 
            // TxtImportFile
            // 
            this.TxtImportFile.Location = new System.Drawing.Point(133, 103);
            this.TxtImportFile.Name = "TxtImportFile";
            this.TxtImportFile.Size = new System.Drawing.Size(362, 20);
            this.TxtImportFile.TabIndex = 1;
            // 
            // CmdBrowse
            // 
            this.CmdBrowse.Location = new System.Drawing.Point(48, 98);
            this.CmdBrowse.Name = "CmdBrowse";
            this.CmdBrowse.Size = new System.Drawing.Size(79, 28);
            this.CmdBrowse.TabIndex = 2;
            this.CmdBrowse.Text = "Browse";
            this.CmdBrowse.UseVisualStyleBackColor = true;
            this.CmdBrowse.Click += new System.EventHandler(this.CmdBrowse_Click);
            // 
            // CmdExit
            // 
            this.CmdExit.Location = new System.Drawing.Point(405, 158);
            this.CmdExit.Name = "CmdExit";
            this.CmdExit.Size = new System.Drawing.Size(79, 28);
            this.CmdExit.TabIndex = 3;
            this.CmdExit.Text = "Exit";
            this.CmdExit.UseVisualStyleBackColor = true;
            this.CmdExit.Click += new System.EventHandler(this.CmdExit_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(64, 42);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(388, 20);
            this.label2.TabIndex = 4;
            this.label2.Text = "(Load Excel Data Into Tableau-Compatible Database)";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(173, 16);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(135, 20);
            this.label3.TabIndex = 5;
            this.label3.Text = "SBS Data Import ";
            // 
            // CmdImport
            // 
            this.CmdImport.Enabled = false;
            this.CmdImport.Location = new System.Drawing.Point(48, 132);
            this.CmdImport.Name = "CmdImport";
            this.CmdImport.Size = new System.Drawing.Size(79, 28);
            this.CmdImport.TabIndex = 6;
            this.CmdImport.Text = "Import Data";
            this.CmdImport.UseVisualStyleBackColor = true;
            this.CmdImport.Click += new System.EventHandler(this.CmdImport_Click);
            // 
            // FrmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(535, 242);
            this.Controls.Add(this.CmdImport);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.CmdExit);
            this.Controls.Add(this.CmdBrowse);
            this.Controls.Add(this.TxtImportFile);
            this.Controls.Add(this.label1);
            this.Name = "FrmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Journal Entry Spreadsheet Import";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox TxtImportFile;
        private System.Windows.Forms.Button CmdBrowse;
        private System.Windows.Forms.Button CmdExit;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button CmdImport;
    }
}


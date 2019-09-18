namespace ManageEmailLists
{
    partial class fMain
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
            this.btnImportExcel = new System.Windows.Forms.Button();
            this.btnExportEmailsWithoutDuplicates = new System.Windows.Forms.Button();
            this.btnExportPrivateBusinessEmails = new System.Windows.Forms.Button();
            this.lbDuplicateEmails = new System.Windows.Forms.ListBox();
            this.btnFindDuplicates = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnImportExcel
            // 
            this.btnImportExcel.Location = new System.Drawing.Point(12, 22);
            this.btnImportExcel.Name = "btnImportExcel";
            this.btnImportExcel.Size = new System.Drawing.Size(163, 23);
            this.btnImportExcel.TabIndex = 0;
            this.btnImportExcel.Text = "Import...";
            this.btnImportExcel.UseVisualStyleBackColor = true;
            this.btnImportExcel.Click += new System.EventHandler(this.BtnImportExcel_Click);
            // 
            // btnExportEmailsWithoutDuplicates
            // 
            this.btnExportEmailsWithoutDuplicates.Enabled = false;
            this.btnExportEmailsWithoutDuplicates.Location = new System.Drawing.Point(12, 80);
            this.btnExportEmailsWithoutDuplicates.Name = "btnExportEmailsWithoutDuplicates";
            this.btnExportEmailsWithoutDuplicates.Size = new System.Drawing.Size(163, 23);
            this.btnExportEmailsWithoutDuplicates.TabIndex = 1;
            this.btnExportEmailsWithoutDuplicates.Text = "Export emails with duplicates";
            this.btnExportEmailsWithoutDuplicates.UseVisualStyleBackColor = true;
            this.btnExportEmailsWithoutDuplicates.Click += new System.EventHandler(this.BtnExportEmailsWithoutDuplicates_Click);
            // 
            // btnExportPrivateBusinessEmails
            // 
            this.btnExportPrivateBusinessEmails.Enabled = false;
            this.btnExportPrivateBusinessEmails.Location = new System.Drawing.Point(12, 109);
            this.btnExportPrivateBusinessEmails.Name = "btnExportPrivateBusinessEmails";
            this.btnExportPrivateBusinessEmails.Size = new System.Drawing.Size(163, 23);
            this.btnExportPrivateBusinessEmails.TabIndex = 2;
            this.btnExportPrivateBusinessEmails.Text = "Export private and business emails";
            this.btnExportPrivateBusinessEmails.UseVisualStyleBackColor = true;
            this.btnExportPrivateBusinessEmails.Click += new System.EventHandler(this.BtnExportPrivateBusinessEmails_Click);
            // 
            // lbDuplicateEmails
            // 
            this.lbDuplicateEmails.FormattingEnabled = true;
            this.lbDuplicateEmails.Location = new System.Drawing.Point(190, 22);
            this.lbDuplicateEmails.Name = "lbDuplicateEmails";
            this.lbDuplicateEmails.Size = new System.Drawing.Size(243, 329);
            this.lbDuplicateEmails.TabIndex = 3;
            // 
            // btnFindDuplicates
            // 
            this.btnFindDuplicates.Enabled = false;
            this.btnFindDuplicates.Location = new System.Drawing.Point(12, 51);
            this.btnFindDuplicates.Name = "btnFindDuplicates";
            this.btnFindDuplicates.Size = new System.Drawing.Size(163, 23);
            this.btnFindDuplicates.TabIndex = 4;
            this.btnFindDuplicates.Text = "Find duplicates";
            this.btnFindDuplicates.UseVisualStyleBackColor = true;
            this.btnFindDuplicates.Click += new System.EventHandler(this.BtnFindDuplicates_Click);
            // 
            // fMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(454, 375);
            this.Controls.Add(this.btnFindDuplicates);
            this.Controls.Add(this.lbDuplicateEmails);
            this.Controls.Add(this.btnExportPrivateBusinessEmails);
            this.Controls.Add(this.btnExportEmailsWithoutDuplicates);
            this.Controls.Add(this.btnImportExcel);
            this.Name = "fMain";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnImportExcel;
        private System.Windows.Forms.Button btnExportEmailsWithoutDuplicates;
        private System.Windows.Forms.Button btnExportPrivateBusinessEmails;
        private System.Windows.Forms.ListBox lbDuplicateEmails;
        private System.Windows.Forms.Button btnFindDuplicates;
    }
}


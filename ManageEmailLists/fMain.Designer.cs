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
            this.btnExportPrivateCorporateEmails = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.btnFindDuplicates = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnImportExcel
            // 
            this.btnImportExcel.Location = new System.Drawing.Point(12, 22);
            this.btnImportExcel.Name = "btnImportExcel";
            this.btnImportExcel.Size = new System.Drawing.Size(93, 23);
            this.btnImportExcel.TabIndex = 0;
            this.btnImportExcel.Text = "Import...";
            this.btnImportExcel.UseVisualStyleBackColor = true;
            this.btnImportExcel.Click += new System.EventHandler(this.BtnImportExcel_Click);
            // 
            // btnExportEmailsWithoutDuplicates
            // 
            this.btnExportEmailsWithoutDuplicates.Location = new System.Drawing.Point(12, 80);
            this.btnExportEmailsWithoutDuplicates.Name = "btnExportEmailsWithoutDuplicates";
            this.btnExportEmailsWithoutDuplicates.Size = new System.Drawing.Size(93, 23);
            this.btnExportEmailsWithoutDuplicates.TabIndex = 1;
            this.btnExportEmailsWithoutDuplicates.Text = "button2";
            this.btnExportEmailsWithoutDuplicates.UseVisualStyleBackColor = true;
            this.btnExportEmailsWithoutDuplicates.Click += new System.EventHandler(this.BtnExportEmailsWithoutDuplicates_Click);
            // 
            // btnExportPrivateCorporateEmails
            // 
            this.btnExportPrivateCorporateEmails.Location = new System.Drawing.Point(12, 109);
            this.btnExportPrivateCorporateEmails.Name = "btnExportPrivateCorporateEmails";
            this.btnExportPrivateCorporateEmails.Size = new System.Drawing.Size(93, 23);
            this.btnExportPrivateCorporateEmails.TabIndex = 2;
            this.btnExportPrivateCorporateEmails.Text = "button3";
            this.btnExportPrivateCorporateEmails.UseVisualStyleBackColor = true;
            this.btnExportPrivateCorporateEmails.Click += new System.EventHandler(this.BtnExportPrivateCorporateEmails_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(126, 22);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(197, 329);
            this.listBox1.TabIndex = 3;
            // 
            // btnFindDuplicates
            // 
            this.btnFindDuplicates.Location = new System.Drawing.Point(12, 51);
            this.btnFindDuplicates.Name = "btnFindDuplicates";
            this.btnFindDuplicates.Size = new System.Drawing.Size(93, 23);
            this.btnFindDuplicates.TabIndex = 4;
            this.btnFindDuplicates.Text = "Find Duplicates";
            this.btnFindDuplicates.UseVisualStyleBackColor = true;
            // 
            // fMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(335, 375);
            this.Controls.Add(this.btnFindDuplicates);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.btnExportPrivateCorporateEmails);
            this.Controls.Add(this.btnExportEmailsWithoutDuplicates);
            this.Controls.Add(this.btnImportExcel);
            this.Name = "fMain";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnImportExcel;
        private System.Windows.Forms.Button btnExportEmailsWithoutDuplicates;
        private System.Windows.Forms.Button btnExportPrivateCorporateEmails;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Button btnFindDuplicates;
    }
}


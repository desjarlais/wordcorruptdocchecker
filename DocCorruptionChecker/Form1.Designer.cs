namespace DocCorruptionChecker
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.tbxFileName = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.btnScanDocument = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.btnCopy = new System.Windows.Forms.Button();
            this.chkRemoveAllFallbackTags = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(24, 17);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(101, 25);
            this.label1.TabIndex = 0;
            this.label1.Text = "File path:";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.DefaultExt = "\".docx\"";
            this.openFileDialog1.Filter = "\"Word Open XML Documents | *.docx; *.dotx; *.docm; *.dotm\"";
            this.openFileDialog1.RestoreDirectory = true;
            this.openFileDialog1.Title = "Select Word Document To Fix";
            // 
            // tbxFileName
            // 
            this.tbxFileName.Location = new System.Drawing.Point(130, 12);
            this.tbxFileName.Margin = new System.Windows.Forms.Padding(6);
            this.tbxFileName.Name = "tbxFileName";
            this.tbxFileName.Size = new System.Drawing.Size(1042, 31);
            this.tbxFileName.TabIndex = 1;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(1188, 8);
            this.btnBrowse.Margin = new System.Windows.Forms.Padding(6);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(88, 44);
            this.btnBrowse.TabIndex = 2;
            this.btnBrowse.Text = "...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // btnScanDocument
            // 
            this.btnScanDocument.Location = new System.Drawing.Point(1288, 8);
            this.btnScanDocument.Margin = new System.Windows.Forms.Padding(6);
            this.btnScanDocument.Name = "btnScanDocument";
            this.btnScanDocument.Size = new System.Drawing.Size(172, 44);
            this.btnScanDocument.TabIndex = 3;
            this.btnScanDocument.Text = "Fix Document";
            this.btnScanDocument.UseVisualStyleBackColor = true;
            this.btnScanDocument.Click += new System.EventHandler(this.btnScanDocument_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.HorizontalScrollbar = true;
            this.listBox1.ItemHeight = 25;
            this.listBox1.Location = new System.Drawing.Point(30, 62);
            this.listBox1.Margin = new System.Windows.Forms.Padding(6);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(1426, 729);
            this.listBox1.TabIndex = 4;
            // 
            // btnCopy
            // 
            this.btnCopy.Location = new System.Drawing.Point(1236, 806);
            this.btnCopy.Margin = new System.Windows.Forms.Padding(6);
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.Size = new System.Drawing.Size(220, 50);
            this.btnCopy.TabIndex = 7;
            this.btnCopy.Text = "Copy Error Details";
            this.btnCopy.UseVisualStyleBackColor = true;
            this.btnCopy.Click += new System.EventHandler(this.btnCopy_Click);
            // 
            // chkRemoveAllFallbackTags
            // 
            this.chkRemoveAllFallbackTags.AutoSize = true;
            this.chkRemoveAllFallbackTags.Location = new System.Drawing.Point(29, 806);
            this.chkRemoveAllFallbackTags.Name = "chkRemoveAllFallbackTags";
            this.chkRemoveAllFallbackTags.Size = new System.Drawing.Size(294, 29);
            this.chkRemoveAllFallbackTags.TabIndex = 8;
            this.chkRemoveAllFallbackTags.Text = "Remove All Fallback Tags";
            this.chkRemoveAllFallbackTags.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(192F, 192F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(1480, 871);
            this.Controls.Add(this.chkRemoveAllFallbackTags);
            this.Controls.Add(this.btnCopy);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.btnScanDocument);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.tbxFileName);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(6);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Word Corrupt Document Checker";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox tbxFileName;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Button btnScanDocument;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Button btnCopy;
        private System.Windows.Forms.CheckBox chkRemoveAllFallbackTags;
    }
}


namespace DocCorruptionChecker
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.label1 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.tbxFileName = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.btnFixDocument = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.btnCopy = new System.Windows.Forms.Button();
            this.chkRemoveAllFallbackTags = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(50, 13);
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
            this.tbxFileName.Location = new System.Drawing.Point(65, 6);
            this.tbxFileName.Name = "tbxFileName";
            this.tbxFileName.Size = new System.Drawing.Size(523, 20);
            this.tbxFileName.TabIndex = 1;
            // 
            // btnBrowse
            // 
            this.btnBrowse.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnBrowse.Location = new System.Drawing.Point(594, 4);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(44, 22);
            this.btnBrowse.TabIndex = 2;
            this.btnBrowse.Text = "...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // btnFixDocument
            // 
            this.btnFixDocument.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnFixDocument.Location = new System.Drawing.Point(644, 4);
            this.btnFixDocument.Name = "btnFixDocument";
            this.btnFixDocument.Size = new System.Drawing.Size(86, 22);
            this.btnFixDocument.TabIndex = 3;
            this.btnFixDocument.Text = "Fix Document";
            this.btnFixDocument.UseVisualStyleBackColor = true;
            this.btnFixDocument.Click += new System.EventHandler(this.btnFixDocument_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.HorizontalScrollbar = true;
            this.listBox1.Location = new System.Drawing.Point(15, 31);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(715, 355);
            this.listBox1.TabIndex = 4;
            // 
            // btnCopy
            // 
            this.btnCopy.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnCopy.Location = new System.Drawing.Point(618, 403);
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.Size = new System.Drawing.Size(110, 25);
            this.btnCopy.TabIndex = 7;
            this.btnCopy.Text = "Copy Error Details";
            this.btnCopy.UseVisualStyleBackColor = true;
            this.btnCopy.Click += new System.EventHandler(this.btnCopy_Click);
            // 
            // chkRemoveAllFallbackTags
            // 
            this.chkRemoveAllFallbackTags.AutoSize = true;
            this.chkRemoveAllFallbackTags.Location = new System.Drawing.Point(14, 403);
            this.chkRemoveAllFallbackTags.Margin = new System.Windows.Forms.Padding(2);
            this.chkRemoveAllFallbackTags.Name = "chkRemoveAllFallbackTags";
            this.chkRemoveAllFallbackTags.Size = new System.Drawing.Size(150, 17);
            this.chkRemoveAllFallbackTags.TabIndex = 8;
            this.chkRemoveAllFallbackTags.Text = "Remove All Fallback Tags";
            this.chkRemoveAllFallbackTags.UseVisualStyleBackColor = true;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(740, 436);
            this.Controls.Add(this.chkRemoveAllFallbackTags);
            this.Controls.Add(this.btnCopy);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.btnFixDocument);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.tbxFileName);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Corrupt Word Document Checker";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox tbxFileName;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Button btnFixDocument;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Button btnCopy;
        private System.Windows.Forms.CheckBox chkRemoveAllFallbackTags;
    }
}


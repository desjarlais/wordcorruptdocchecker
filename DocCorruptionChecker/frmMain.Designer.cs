namespace DocCorruptionChecker
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmMain));
            this.label1 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.tbxFileName = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.btnFixDocument = new System.Windows.Forms.Button();
            this.lstOutput = new System.Windows.Forms.ListBox();
            this.btnCopy = new System.Windows.Forms.Button();
            this.BtnSettings = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(24, 16);
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
            this.btnBrowse.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnBrowse.Location = new System.Drawing.Point(1188, 8);
            this.btnBrowse.Margin = new System.Windows.Forms.Padding(6);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(88, 44);
            this.btnBrowse.TabIndex = 2;
            this.btnBrowse.Text = "...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.BtnBrowse_Click);
            // 
            // btnFixDocument
            // 
            this.btnFixDocument.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnFixDocument.Location = new System.Drawing.Point(1288, 8);
            this.btnFixDocument.Margin = new System.Windows.Forms.Padding(6);
            this.btnFixDocument.Name = "btnFixDocument";
            this.btnFixDocument.Size = new System.Drawing.Size(172, 44);
            this.btnFixDocument.TabIndex = 3;
            this.btnFixDocument.Text = "Fix Document";
            this.btnFixDocument.UseVisualStyleBackColor = true;
            this.btnFixDocument.Click += new System.EventHandler(this.BtnFixDocument_Click);
            // 
            // lstOutput
            // 
            this.lstOutput.FormattingEnabled = true;
            this.lstOutput.HorizontalScrollbar = true;
            this.lstOutput.ItemHeight = 25;
            this.lstOutput.Location = new System.Drawing.Point(30, 62);
            this.lstOutput.Margin = new System.Windows.Forms.Padding(6);
            this.lstOutput.Name = "lstOutput";
            this.lstOutput.Size = new System.Drawing.Size(1426, 704);
            this.lstOutput.TabIndex = 4;
            // 
            // btnCopy
            // 
            this.btnCopy.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnCopy.Location = new System.Drawing.Point(1240, 777);
            this.btnCopy.Margin = new System.Windows.Forms.Padding(6);
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.Size = new System.Drawing.Size(220, 50);
            this.btnCopy.TabIndex = 7;
            this.btnCopy.Text = "Copy Error Details";
            this.btnCopy.UseVisualStyleBackColor = true;
            this.btnCopy.Click += new System.EventHandler(this.BtnCopy_Click);
            // 
            // BtnSettings
            // 
            this.BtnSettings.Location = new System.Drawing.Point(1080, 775);
            this.BtnSettings.Name = "BtnSettings";
            this.BtnSettings.Size = new System.Drawing.Size(147, 50);
            this.BtnSettings.TabIndex = 9;
            this.BtnSettings.Text = "Settings";
            this.BtnSettings.UseVisualStyleBackColor = true;
            this.BtnSettings.Click += new System.EventHandler(this.BtnSettings_Click);
            // 
            // FrmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(192F, 192F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(1480, 848);
            this.Controls.Add(this.BtnSettings);
            this.Controls.Add(this.btnCopy);
            this.Controls.Add(this.lstOutput);
            this.Controls.Add(this.btnFixDocument);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.tbxFileName);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(6);
            this.MaximizeBox = false;
            this.Name = "FrmMain";
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
        private System.Windows.Forms.ListBox lstOutput;
        private System.Windows.Forms.Button btnCopy;
        private System.Windows.Forms.Button BtnSettings;
    }
}


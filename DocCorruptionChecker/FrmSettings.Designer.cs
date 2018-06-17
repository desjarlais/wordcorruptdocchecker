namespace DocCorruptionChecker
{
    partial class FrmSettings
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmSettings));
            this.ckRemoveFallback = new System.Windows.Forms.CheckBox();
            this.ckOpenInWord = new System.Windows.Forms.CheckBox();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.lblDisplayVersion = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // ckRemoveFallback
            // 
            this.ckRemoveFallback.AutoSize = true;
            this.ckRemoveFallback.Location = new System.Drawing.Point(16, 17);
            this.ckRemoveFallback.Name = "ckRemoveFallback";
            this.ckRemoveFallback.Size = new System.Drawing.Size(217, 24);
            this.ckRemoveFallback.TabIndex = 0;
            this.ckRemoveFallback.Text = "Remove All Fallback Tags";
            this.ckRemoveFallback.UseVisualStyleBackColor = true;
            // 
            // ckOpenInWord
            // 
            this.ckOpenInWord.AutoSize = true;
            this.ckOpenInWord.Location = new System.Drawing.Point(16, 49);
            this.ckOpenInWord.Name = "ckOpenInWord";
            this.ckOpenInWord.Size = new System.Drawing.Size(251, 24);
            this.ckOpenInWord.TabIndex = 1;
            this.ckOpenInWord.Text = "Open File in Word After Repair";
            this.ckOpenInWord.UseVisualStyleBackColor = true;
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(161, 124);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 37);
            this.btnOk.TabIndex = 2;
            this.btnOk.Text = "Ok";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.BtnOk_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(242, 124);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(81, 37);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(16, 89);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(117, 20);
            this.label1.TabIndex = 4;
            this.label1.Text = "App Version: ";
            // 
            // lblDisplayVersion
            // 
            this.lblDisplayVersion.AutoSize = true;
            this.lblDisplayVersion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDisplayVersion.Location = new System.Drawing.Point(127, 89);
            this.lblDisplayVersion.Name = "lblDisplayVersion";
            this.lblDisplayVersion.Size = new System.Drawing.Size(0, 20);
            this.lblDisplayVersion.TabIndex = 5;
            // 
            // FrmSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(335, 175);
            this.Controls.Add(this.lblDisplayVersion);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.ckOpenInWord);
            this.Controls.Add(this.ckRemoveFallback);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "FrmSettings";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Settings";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox ckRemoveFallback;
        private System.Windows.Forms.CheckBox ckOpenInWord;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblDisplayVersion;
    }
}
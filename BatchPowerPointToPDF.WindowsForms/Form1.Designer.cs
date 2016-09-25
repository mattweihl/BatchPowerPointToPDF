namespace BatchPowerPointToPDF.WindowsForms
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
            this.officeInstalledLabel = new System.Windows.Forms.Label();
            this.openPPTXBtn = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // officeInstalledLabel
            // 
            this.officeInstalledLabel.AutoSize = true;
            this.officeInstalledLabel.Location = new System.Drawing.Point(54, 9);
            this.officeInstalledLabel.Name = "officeInstalledLabel";
            this.officeInstalledLabel.Size = new System.Drawing.Size(139, 25);
            this.officeInstalledLabel.TabIndex = 0;
            this.officeInstalledLabel.Text = "Office Installed: ";
            // 
            // openPPTXBtn
            // 
            this.openPPTXBtn.Location = new System.Drawing.Point(12, 37);
            this.openPPTXBtn.Name = "openPPTXBtn";
            this.openPPTXBtn.Size = new System.Drawing.Size(215, 36);
            this.openPPTXBtn.TabIndex = 1;
            this.openPPTXBtn.Text = "Open PowerPoint File(s)";
            this.openPPTXBtn.UseVisualStyleBackColor = true;
            this.openPPTXBtn.Click += new System.EventHandler(this.openPPTXBtn_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 79);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(215, 36);
            this.button1.TabIndex = 3;
            this.button1.Text = "Export To PDF";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(239, 152);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.openPPTXBtn);
            this.Controls.Add(this.officeInstalledLabel);
            this.DoubleBuffered = true;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Batch PowerPoint To PDF Tool";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label officeInstalledLabel;
        private System.Windows.Forms.Button openPPTXBtn;
        private System.Windows.Forms.Button button1;
    }
}


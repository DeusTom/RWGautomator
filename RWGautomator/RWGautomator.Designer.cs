namespace KRGautomator
{
    partial class RWGautomator
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.uploadDocument = new System.Windows.Forms.Button();
            this.generateRoots = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // uploadDocument
            // 
            this.uploadDocument.Location = new System.Drawing.Point(12, 12);
            this.uploadDocument.Name = "uploadDocument";
            this.uploadDocument.Size = new System.Drawing.Size(247, 31);
            this.uploadDocument.TabIndex = 0;
            this.uploadDocument.Text = "Upload document";
            this.uploadDocument.UseVisualStyleBackColor = true;
            this.uploadDocument.Click += new System.EventHandler(this.uploadDocument_Click);
            // 
            // generateRoots
            // 
            this.generateRoots.Location = new System.Drawing.Point(12, 49);
            this.generateRoots.Name = "generateRoots";
            this.generateRoots.Size = new System.Drawing.Size(247, 41);
            this.generateRoots.TabIndex = 1;
            this.generateRoots.Text = "Generate roots";
            this.generateRoots.UseVisualStyleBackColor = true;
            this.generateRoots.Click += new System.EventHandler(this.generateRoots_Click);
            // 
            // KRGautomator
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(271, 102);
            this.Controls.Add(this.generateRoots);
            this.Controls.Add(this.uploadDocument);
            this.Name = "KRGautomator";
            this.Text = "RWG automator";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.KRGautomator_FormClosing);
            this.ResumeLayout(false);

        }

        #endregion

        private Button uploadDocument;
        private Button generateRoots;
    }
}
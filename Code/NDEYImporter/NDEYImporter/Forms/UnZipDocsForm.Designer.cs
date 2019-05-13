namespace NDEYImporter.Forms
{
    partial class UnZipDocsForm
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
            this.tvSubDirs = new System.Windows.Forms.TreeView();
            this.SuspendLayout();
            // 
            // tvSubDirs
            // 
            this.tvSubDirs.CheckBoxes = true;
            this.tvSubDirs.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tvSubDirs.FullRowSelect = true;
            this.tvSubDirs.Location = new System.Drawing.Point(0, 0);
            this.tvSubDirs.Name = "tvSubDirs";
            this.tvSubDirs.Size = new System.Drawing.Size(292, 494);
            this.tvSubDirs.TabIndex = 1;
            // 
            // UnZipDocsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(292, 494);
            this.Controls.Add(this.tvSubDirs);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "UnZipDocsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "请选择要解压的项目!";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TreeView tvSubDirs;
    }
}
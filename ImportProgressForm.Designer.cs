namespace renamePDF_v2
{
    partial class ImportProgressForm
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
            lblTitle = new Label();
            lblFile = new Label();
            lblCount = new Label();
            progressBar = new ProgressBar();
            SuspendLayout();
            // 
            // lblTitle
            // 
            lblTitle.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            lblTitle.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            lblTitle.Location = new Point(12, 9);
            lblTitle.Name = "lblTitle";
            lblTitle.Size = new Size(290, 29);
            lblTitle.TabIndex = 0;
            lblTitle.Text = "label1";
            lblTitle.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // lblFile
            // 
            lblFile.Location = new Point(12, 58);
            lblFile.Name = "lblFile";
            lblFile.Size = new Size(290, 26);
            lblFile.TabIndex = 1;
            lblFile.Text = "label1";
            lblFile.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // lblCount
            // 
            lblCount.Location = new Point(12, 96);
            lblCount.Name = "lblCount";
            lblCount.Size = new Size(290, 18);
            lblCount.TabIndex = 2;
            lblCount.Text = "0/0";
            lblCount.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // progressBar
            // 
            progressBar.Location = new Point(12, 131);
            progressBar.Name = "progressBar";
            progressBar.Size = new Size(290, 23);
            progressBar.TabIndex = 3;
            // 
            // ImportProgressForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(314, 186);
            Controls.Add(progressBar);
            Controls.Add(lblCount);
            Controls.Add(lblFile);
            Controls.Add(lblTitle);
            Name = "ImportProgressForm";
            StartPosition = FormStartPosition.CenterScreen;
            ResumeLayout(false);
        }

        #endregion

        private Label lblTitle;
        private Label lblFile;
        private Label lblCount;
        private ProgressBar progressBar;
    }
}
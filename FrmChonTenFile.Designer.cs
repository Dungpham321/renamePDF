using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace renamePDF_v2
{
    partial class FrmChonTenFile
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }

            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            lblTimKiem = new Label();
            txtTimKiem = new TextBox();
            cboTenFile = new ComboBox();
            btnChon = new Button();
            btnHuy = new Button();
            lblTenFile = new Label();

            SuspendLayout();

            // 
            // lblTimKiem
            // 
            lblTimKiem.AutoSize = true;
            lblTimKiem.Location = new Point(20, 20);
            lblTimKiem.Name = "lblTimKiem";
            lblTimKiem.Size = new Size(70, 15);
            lblTimKiem.TabIndex = 0;
            lblTimKiem.Text = "Tìm kiếm:";

            // 
            // txtTimKiem
            // 
            txtTimKiem.Location = new Point(105, 17);
            txtTimKiem.Name = "txtTimKiem";
            txtTimKiem.PlaceholderText = "Nhập tên cần tìm...";
            txtTimKiem.Size = new Size(450, 23);
            txtTimKiem.TabIndex = 1;

            // 
            // lblTenFile
            // 
            lblTenFile.AutoSize = true;
            lblTenFile.Location = new Point(20, 58);
            lblTenFile.Name = "lblTenFile";
            lblTenFile.Size = new Size(55, 15);
            lblTenFile.TabIndex = 2;
            lblTenFile.Text = "Tên file:";

            // 
            // cboTenFile
            // 
            cboTenFile.DropDownStyle = ComboBoxStyle.DropDown;
            cboTenFile.FormattingEnabled = true;
            cboTenFile.Location = new Point(105, 55);
            cboTenFile.Name = "cboTenFile";
            cboTenFile.Size = new Size(450, 23);
            cboTenFile.TabIndex = 3;

            // 
            // btnChon
            // 
            btnChon.Location = new Point(370, 100);
            btnChon.Name = "btnChon";
            btnChon.Size = new Size(85, 30);
            btnChon.TabIndex = 4;
            btnChon.Text = "Chọn";
            btnChon.UseVisualStyleBackColor = true;

            // 
            // btnHuy
            // 
            btnHuy.DialogResult = DialogResult.Cancel;
            btnHuy.Location = new Point(470, 100);
            btnHuy.Name = "btnHuy";
            btnHuy.Size = new Size(85, 30);
            btnHuy.TabIndex = 5;
            btnHuy.Text = "Hủy";
            btnHuy.UseVisualStyleBackColor = true;

            // 
            // FrmChonTenFile
            // 
            AcceptButton = btnChon;
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            CancelButton = btnHuy;
            ClientSize = new Size(580, 150);
            Controls.Add(btnHuy);
            Controls.Add(btnChon);
            Controls.Add(cboTenFile);
            Controls.Add(lblTenFile);
            Controls.Add(txtTimKiem);
            Controls.Add(lblTimKiem);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "FrmChonTenFile";
            StartPosition = FormStartPosition.CenterParent;
            Text = "Chọn tên PDF";

            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label lblTimKiem;
        private TextBox txtTimKiem;
        private ComboBox cboTenFile;
        private Button btnChon;
        private Button btnHuy;
        private Label lblTenFile;
    }
}
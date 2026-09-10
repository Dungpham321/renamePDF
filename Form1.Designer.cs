namespace renamePDF_v2
{
    partial class Form1
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
            tabControl1 = new TabControl();
            tabPage1 = new TabPage();
            btnXuLyDuLieu = new Button();
            btnSave = new Button();
            btnNhanDang = new Button();
            txtSearch = new TextBox();
            flowPdf = new FlowLayoutPanel();
            panel1 = new Panel();
            btnImport = new Button();
            btnchonfile = new Button();
            label1 = new Label();
            txtPathFile = new TextBox();
            treeView1 = new TreeView();
            tabPage2 = new TabPage();
            btnChonDanhMuc = new Button();
            btnChonThuMuc = new Button();
            txturlapi = new TextBox();
            txtpathdanhmuc = new TextBox();
            txtpaththumuc = new TextBox();
            btnSaveConfig = new Button();
            label4 = new Label();
            label3 = new Label();
            label2 = new Label();
            panel2 = new Panel();
            dataGridView = new DataGridView();
            tabControl1.SuspendLayout();
            tabPage1.SuspendLayout();
            panel1.SuspendLayout();
            tabPage2.SuspendLayout();
            panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView).BeginInit();
            SuspendLayout();
            // 
            // tabControl1
            // 
            tabControl1.Controls.Add(tabPage1);
            tabControl1.Controls.Add(tabPage2);
            tabControl1.Dock = DockStyle.Fill;
            tabControl1.Location = new Point(0, 0);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new Size(1246, 684);
            tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            tabPage1.AutoScroll = true;
            tabPage1.Controls.Add(btnXuLyDuLieu);
            tabPage1.Controls.Add(btnSave);
            tabPage1.Controls.Add(btnNhanDang);
            tabPage1.Controls.Add(txtSearch);
            tabPage1.Controls.Add(flowPdf);
            tabPage1.Controls.Add(panel1);
            tabPage1.Controls.Add(treeView1);
            tabPage1.Location = new Point(4, 24);
            tabPage1.Name = "tabPage1";
            tabPage1.Padding = new Padding(3);
            tabPage1.Size = new Size(1238, 656);
            tabPage1.TabIndex = 0;
            tabPage1.Text = "Kiểm tra dữ liệu";
            tabPage1.UseVisualStyleBackColor = true;
            // 
            // btnXuLyDuLieu
            // 
            btnXuLyDuLieu.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnXuLyDuLieu.Location = new Point(989, 34);
            btnXuLyDuLieu.Name = "btnXuLyDuLieu";
            btnXuLyDuLieu.Size = new Size(105, 33);
            btnXuLyDuLieu.TabIndex = 11;
            btnXuLyDuLieu.Text = "Xử lý dữ liệu";
            btnXuLyDuLieu.UseVisualStyleBackColor = true;
            btnXuLyDuLieu.Click += btnXuLyDuLieu_Click;
            // 
            // btnSave
            // 
            btnSave.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnSave.BackColor = Color.MediumSpringGreen;
            btnSave.ForeColor = SystemColors.ActiveCaptionText;
            btnSave.Location = new Point(1100, 34);
            btnSave.Name = "btnSave";
            btnSave.Size = new Size(135, 32);
            btnSave.TabIndex = 10;
            btnSave.Text = "Cập nhật tên file";
            btnSave.UseVisualStyleBackColor = false;
            btnSave.Click += btnSave_Click;
            // 
            // btnNhanDang
            // 
            btnNhanDang.Location = new Point(348, 34);
            btnNhanDang.Name = "btnNhanDang";
            btnNhanDang.Size = new Size(75, 33);
            btnNhanDang.TabIndex = 9;
            btnNhanDang.Text = "Nhận dạng";
            btnNhanDang.UseVisualStyleBackColor = true;
            btnNhanDang.Click += btnNhanDang_Click;
            // 
            // txtSearch
            // 
            txtSearch.BorderStyle = BorderStyle.FixedSingle;
            txtSearch.Location = new Point(8, 72);
            txtSearch.Name = "txtSearch";
            txtSearch.PlaceholderText = "Tìm kiếm";
            txtSearch.Size = new Size(334, 23);
            txtSearch.TabIndex = 8;
            // 
            // flowPdf
            // 
            flowPdf.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            flowPdf.AutoScroll = true;
            flowPdf.BorderStyle = BorderStyle.FixedSingle;
            flowPdf.Location = new Point(348, 72);
            flowPdf.Name = "flowPdf";
            flowPdf.Size = new Size(887, 581);
            flowPdf.TabIndex = 7;
            // 
            // panel1
            // 
            panel1.BorderStyle = BorderStyle.FixedSingle;
            panel1.Controls.Add(btnImport);
            panel1.Controls.Add(btnchonfile);
            panel1.Controls.Add(label1);
            panel1.Controls.Add(txtPathFile);
            panel1.Location = new Point(8, 6);
            panel1.Name = "panel1";
            panel1.Padding = new Padding(1);
            panel1.Size = new Size(334, 60);
            panel1.TabIndex = 6;
            // 
            // btnImport
            // 
            btnImport.Enabled = false;
            btnImport.Location = new Point(263, 25);
            btnImport.Name = "btnImport";
            btnImport.Size = new Size(62, 26);
            btnImport.TabIndex = 4;
            btnImport.Text = "Import";
            btnImport.UseVisualStyleBackColor = true;
            btnImport.Click += btnImport_Click_1;
            // 
            // btnchonfile
            // 
            btnchonfile.Location = new Point(184, 25);
            btnchonfile.Name = "btnchonfile";
            btnchonfile.Size = new Size(75, 26);
            btnchonfile.TabIndex = 3;
            btnchonfile.Text = "Chọn file";
            btnchonfile.UseVisualStyleBackColor = true;
            btnchonfile.Click += btnchonfile_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            label1.Location = new Point(5, 5);
            label1.Name = "label1";
            label1.Size = new Size(106, 15);
            label1.TabIndex = 1;
            label1.Text = "Import File Excell:";
            // 
            // txtPathFile
            // 
            txtPathFile.Location = new Point(3, 26);
            txtPathFile.Name = "txtPathFile";
            txtPathFile.ReadOnly = true;
            txtPathFile.Size = new Size(177, 23);
            txtPathFile.TabIndex = 2;
            // 
            // treeView1
            // 
            treeView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;
            treeView1.Location = new Point(8, 101);
            treeView1.Name = "treeView1";
            treeView1.Size = new Size(334, 552);
            treeView1.TabIndex = 5;
            // 
            // tabPage2
            // 
            tabPage2.Controls.Add(btnChonDanhMuc);
            tabPage2.Controls.Add(btnChonThuMuc);
            tabPage2.Controls.Add(txturlapi);
            tabPage2.Controls.Add(txtpathdanhmuc);
            tabPage2.Controls.Add(txtpaththumuc);
            tabPage2.Controls.Add(btnSaveConfig);
            tabPage2.Controls.Add(label4);
            tabPage2.Controls.Add(label3);
            tabPage2.Controls.Add(label2);
            tabPage2.Controls.Add(panel2);
            tabPage2.Location = new Point(4, 24);
            tabPage2.Name = "tabPage2";
            tabPage2.Padding = new Padding(3);
            tabPage2.Size = new Size(1238, 656);
            tabPage2.TabIndex = 1;
            tabPage2.Text = "Cấu hình";
            tabPage2.UseVisualStyleBackColor = true;
            // 
            // btnChonDanhMuc
            // 
            btnChonDanhMuc.Location = new Point(433, 58);
            btnChonDanhMuc.Name = "btnChonDanhMuc";
            btnChonDanhMuc.Size = new Size(122, 26);
            btnChonDanhMuc.TabIndex = 9;
            btnChonDanhMuc.Text = "Chọn danh mục";
            btnChonDanhMuc.UseVisualStyleBackColor = true;
            btnChonDanhMuc.Click += btnChonDanhMuc_Click;
            // 
            // btnChonThuMuc
            // 
            btnChonThuMuc.Location = new Point(433, 13);
            btnChonThuMuc.Name = "btnChonThuMuc";
            btnChonThuMuc.Size = new Size(99, 26);
            btnChonThuMuc.TabIndex = 8;
            btnChonThuMuc.Text = "Chọn thư mục";
            btnChonThuMuc.UseVisualStyleBackColor = true;
            btnChonThuMuc.Click += btnChonThuMuc_Click;
            // 
            // txturlapi
            // 
            txturlapi.Location = new Point(137, 111);
            txturlapi.Name = "txturlapi";
            txturlapi.Size = new Size(295, 23);
            txturlapi.TabIndex = 7;
            // 
            // txtpathdanhmuc
            // 
            txtpathdanhmuc.Location = new Point(137, 59);
            txtpathdanhmuc.Name = "txtpathdanhmuc";
            txtpathdanhmuc.ReadOnly = true;
            txtpathdanhmuc.Size = new Size(295, 23);
            txtpathdanhmuc.TabIndex = 6;
            // 
            // txtpaththumuc
            // 
            txtpaththumuc.Location = new Point(137, 14);
            txtpaththumuc.Name = "txtpaththumuc";
            txtpaththumuc.ReadOnly = true;
            txtpaththumuc.Size = new Size(295, 23);
            txtpaththumuc.TabIndex = 5;
            // 
            // btnSaveConfig
            // 
            btnSaveConfig.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnSaveConfig.Location = new Point(466, 615);
            btnSaveConfig.Name = "btnSaveConfig";
            btnSaveConfig.Size = new Size(89, 33);
            btnSaveConfig.TabIndex = 4;
            btnSaveConfig.Text = "Lưu";
            btnSaveConfig.UseVisualStyleBackColor = true;
            btnSaveConfig.Click += btnSaveConfig_Click;
            // 
            // label4
            // 
            label4.Location = new Point(11, 109);
            label4.Name = "label4";
            label4.Size = new Size(120, 24);
            label4.TabIndex = 3;
            label4.Text = "API AI:";
            label4.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // label3
            // 
            label3.Location = new Point(11, 59);
            label3.Name = "label3";
            label3.Size = new Size(120, 24);
            label3.TabIndex = 2;
            label3.Text = "Danh mục tài liệu:";
            label3.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // label2
            // 
            label2.Location = new Point(11, 14);
            label2.Name = "label2";
            label2.Size = new Size(120, 24);
            label2.TabIndex = 1;
            label2.Text = "Đường dẫn thư mục:";
            label2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // panel2
            // 
            panel2.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            panel2.BorderStyle = BorderStyle.FixedSingle;
            panel2.Controls.Add(dataGridView);
            panel2.Location = new Point(570, 6);
            panel2.Name = "panel2";
            panel2.Size = new Size(662, 642);
            panel2.TabIndex = 0;
            // 
            // dataGridView
            // 
            dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView.Dock = DockStyle.Fill;
            dataGridView.Location = new Point(0, 0);
            dataGridView.MultiSelect = false;
            dataGridView.Name = "dataGridView";
            dataGridView.ReadOnly = true;
            dataGridView.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            dataGridView.Size = new Size(660, 640);
            dataGridView.TabIndex = 0;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1246, 684);
            Controls.Add(tabControl1);
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Chuyển đổi tên file";
            WindowState = FormWindowState.Maximized;
            tabControl1.ResumeLayout(false);
            tabPage1.ResumeLayout(false);
            tabPage1.PerformLayout();
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            tabPage2.ResumeLayout(false);
            tabPage2.PerformLayout();
            panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridView).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private TabControl tabControl1;
        private TabPage tabPage1;
        private TabPage tabPage2;
        private TextBox txtPathFile;
        private Label label1;
        private Button btnchonfile;
        private Button btnImport;
        private TreeView treeView1;
        private Panel panel1;
        private FlowLayoutPanel flowPdf;
        private TextBox txtSearch;
        private Button btnNhanDang;
        private Button btnSave;
        private Panel panel2;
        private Label label3;
        private Label label2;
        private Button btnSaveConfig;
        private Label label4;
        private TextBox txturlapi;
        private TextBox txtpathdanhmuc;
        private TextBox txtpaththumuc;
        private Button btnChonThuMuc;
        private Button btnChonDanhMuc;
        private DataGridView dataGridView;
        private Button btnXuLyDuLieu;
    }
}

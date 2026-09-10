using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using Newtonsoft.Json;
using PDFtoImage;
using renamePDF_v2.Models;
using renamePDF_v2.Services;
using SkiaSharp;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Formats.Tar;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;


namespace renamePDF_v2
{
    public partial class Form1 : Form
    {
        private ExcelImporter _excelImporter = new ExcelImporter();
        string selectedExcelFile = "";
        private class PdfCardInfo
        {
            public Panel Panel { get; set; } = null!;

            public PictureBox PictureBox { get; set; } = null!;

            public Label FileNameLabel { get; set; } = null!;

            public Label StatusLabel { get; set; } = null!;

            public string Path { get; set; } = "";

            public string FileName { get; set; } = "";

            public Dictionary<string, object?> Record { get; set; }
                = null!;

            public bool Loaded { get; set; }

            public bool Failed { get; set; }
        }
        private CancellationTokenSource? _pdfLoadCts;
        private readonly SemaphoreSlim _pdfRenderSemaphore = new SemaphoreSlim(4, 4);
        private readonly ConcurrentDictionary<string, byte> _loadingThumbnails = new ConcurrentDictionary<string, byte>();
        private readonly Dictionary<string, PdfCardInfo> _pdfCards = new Dictionary<string, PdfCardInfo>(StringComparer.OrdinalIgnoreCase);

        private List<Dictionary<string, object?>> _allRecords = new List<Dictionary<string, object?>>();
        private readonly string metadataJsonPath = Path.Combine(Application.StartupPath, "data", "metadata.json");
        private readonly string danhMucJsonPath = Path.Combine(Application.StartupPath, "data", "danhmuctailieu.json");
        private readonly string settingsJsonPath = Path.Combine(Application.StartupPath, "data", "settings.json");

        private bool _loadingPdfList = false;
        private List<Dictionary<string, object?>> danhMucTaiLieu = new();
       

        private List<Dictionary<string, object?>> selectedDangVienRecords = new();
        private DeepSeekClient? _deepSeekClient;
        public Form1()
        {
            InitializeComponent();
            treeView1.AfterSelect += treeView1_AfterSelect;
            txtSearch.TextChanged += txtSearch_TextChanged;
            if (File.Exists(metadataJsonPath))
            {
                LoadTreeFromJson(metadataJsonPath);
            }
            LoadDanhMucVaoDataGridView();
            LoadCauHinh();
            dataGridView.AllowUserToAddRows = false;
            dataGridView.ReadOnly = true;
            btnNhanDang.Enabled = false;
            dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _deepSeekClient = new DeepSeekClient("key dấu rồi");
        }

        private void btnchonfile_Click(object sender, EventArgs e)
        {

            using OpenFileDialog dialog = new OpenFileDialog();

            dialog.Title = "Chọn file Excel";
            dialog.Filter =
                "Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm";

            if (dialog.ShowDialog() != DialogResult.OK)
                return;

            selectedExcelFile = dialog.FileName;

            // Hiển thị đường dẫn
            txtPathFile.Text = selectedExcelFile;

            // Cho phép Import
            btnImport.Enabled = true;

            MessageBox.Show(
                $"Chọn file thành công!",
                "Thành công",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }

        private void btnImport_Click_1(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(selectedExcelFile))
            {
                MessageBox.Show(
                    "Vui lòng chọn file Excel trước.",
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            if (!File.Exists(selectedExcelFile))
            {
                MessageBox.Show(
                    "File Excel không tồn tại.",
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                return;
            }

            try
            {
                btnImport.Enabled = false;
                btnchonfile.Enabled = false;

                Cursor = Cursors.WaitCursor;

                Console.WriteLine("Đang đọc dữ liệu Excel...");

                using (ImportProgressForm progressForm = new ImportProgressForm())
                {
                    progressForm.Show(this);
                    Application.DoEvents();
                    try
                    {
                        var data = _excelImporter.ImportExcel(
                            selectedExcelFile,
                            (current, total, fileName) =>
                            {
                                progressForm.UpdateProgress(
                                    current,
                                    total,
                                    fileName);

                                Application.DoEvents();
                            });

                        // Phần xử lý data tiếp theo của bạn
                        if (data.Count == 0)
                        {
                            MessageBox.Show(
                                "File Excel không có dữ liệu.",
                                "Thông báo",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);

                            return;
                        }
                        string dataFolder = Path.Combine(AppContext.BaseDirectory, "data");
                        // Tạo thư mục data nếu chưa tồn tại
                        Directory.CreateDirectory(dataFolder);

                        string jsonPath = Path.Combine(dataFolder, "metadata" + ".json");

                        Console.WriteLine("Đang tạo file JSON...");

                        string json = JsonConvert.SerializeObject(data, Newtonsoft.Json.Formatting.Indented);

                        File.WriteAllText(
                            jsonPath,
                            json,
                            System.Text.Encoding.UTF8);
                        LoadTreeFromJson(jsonPath);
                        MessageBox.Show(
                            $"Import thành công!\n\n" +
                            $"Số bản ghi: {data.Count:N0}\n\n" +
                            $"File JSON:\n{jsonPath}",
                            "Thành công",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    finally
                    {
                        progressForm.Close();
                    }
                }


                //List<Dictionary<string, object?>> data = ImportExcel(selectedExcelFile);


            }
            catch (Exception ex)
            {

                MessageBox.Show(
                    "Có lỗi khi import:\n\n" + ex.Message,
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                selectedExcelFile = "";
                Cursor = Cursors.Default;

                btnchonfile.Enabled = true;

                btnImport.Enabled = !string.IsNullOrWhiteSpace(selectedExcelFile);
            }
        }
       
        private string GetValue(Dictionary<string, object?> item, string key)
        {
            if (!item.TryGetValue(
                    key,
                    out object? value))
            {
                return "";
            }

            return value?
                .ToString()?
                .Trim()
                ?? "";
        }
        private void LoadTreeFromJson(string jsonPath)
        {
            if (!File.Exists(jsonPath))
            {
                MessageBox.Show(
                    "Không tìm thấy file JSON:\n" + jsonPath,
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            try
            {
                string json = File.ReadAllText(jsonPath, System.Text.Encoding.UTF8);

                var records = JsonConvert.DeserializeObject<List<Dictionary<string, object?>>>(json) ?? new List<Dictionary<string, object?>>();

                _allRecords = records;

                HienThiTree(records);
            }
            catch (Exception ex)
            {
                treeView1.EndUpdate();

                MessageBox.Show(
                    "Lỗi đọc JSON:\n\n" + ex.Message,
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        private void treeView1_AfterSelect(object? sender, TreeViewEventArgs e)
        {
            if (e.Node.Level != 1)
            {
                selectedDangVienRecords.Clear();
                btnNhanDang.Enabled = false;
                return;
            }

            if (e.Node.Tag is List<Dictionary<string, object?>> records)
            {
                selectedDangVienRecords = records;

                btnNhanDang.Enabled = false;///selectedDangVienRecords.Count > 0;

                HienThiDanhSachPdf(records);
            }
            else
            {
                selectedDangVienRecords.Clear();
                btnNhanDang.Enabled = false;
            }
        }

        private async void HienThiDanhSachPdf(List<Dictionary<string, object?>> records)
        {
            _pdfLoadCts?.Cancel();
            _pdfLoadCts?.Dispose();

            _pdfLoadCts =
                new CancellationTokenSource();

            CancellationToken token =
                _pdfLoadCts.Token;

            flowPdf.SuspendLayout();

            try
            {
                foreach (Control control in flowPdf.Controls)
                {
                    DisposeImages(control);
                    control.Dispose();
                }

                flowPdf.Controls.Clear();
                _pdfCards.Clear();
                if (records != null && records.Count > 0)
                {
                    Panel panelCCCD = TaoPanelCCCD(records);
                    flowPdf.Controls.Add(panelCCCD);
                }

                foreach (Dictionary<string, object?> record in records)
                {
                    token.ThrowIfCancellationRequested();

                    string path = GetValue(record, "duong_dan");
                    if (string.IsNullOrWhiteSpace(path))
                        path = GetValue(record, "FOLDER");
                    if (string.IsNullOrWhiteSpace(path))
                        continue;
                    path = path.Trim();
                    string fileName = GetValue(record, "TenFile"); //Path.GetFileName(path);

                    if (string.IsNullOrWhiteSpace(fileName)) continue;

                    PdfCardInfo card = TaoPdfItem(record, path, fileName);

                    flowPdf.Controls.Add(
                        card.Panel);

                    _pdfCards[path] = card;
                }
            }
            catch (OperationCanceledException)
            {
                return;
            }
            finally
            {
                flowPdf.ResumeLayout(true);
            }

            // Load các ảnh đang nhìn thấy
            await LoadVisibleThumbnailsAsync(token);

            // QUAN TRỌNG:
            // Load luôn 10 ảnh cuối
            await LoadLastCardsAsync(token);
        }
        private Panel TaoPanelCCCD(List<Dictionary<string, object?>> records)
        {
            Panel panel = new Panel
            {
                Width = flowPdf.ClientSize.Width - 25,
                Height = 60,
                Margin = new Padding(5),
                Padding = new Padding(10),
                BorderStyle = BorderStyle.FixedSingle
            };

            Label lbl = new Label
            {
                Text = "Căn cước công dân:",
                AutoSize = true,
                Location = new Point(10, 18),
                Font = new Font(
                    "Segoe UI",
                    9F,
                    FontStyle.Bold)
            };

            TextBox txt = new TextBox
            {
                Location = new Point(150, 13),
                Width = 220,
                MaxLength = 12
            };

            Button btn = new Button
            {
                Text = "Lưu",
                Location = new Point(380, 11),
                Width = 70,
                Height = 27
            };

            // Lấy CCCD hiện tại
            string cccd = GetValue(records[0], "CCCD");
            txt.Text = cccd;
            btn.Click += (s, e) =>
            {
                string value = txt.Text.Trim();
                if (value.Length != 12 || !value.All(char.IsDigit))
                {
                    MessageBox.Show(
                        "Số Căn cước công dân phải gồm 12 chữ số.",
                        "Thông báo",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    txt.Focus();
                    return;
                }

                foreach (var record in records)
                {
                    record["CCCD"] = value;
                }

                bool daLuu = LuuCapNhatRecordsVaoJson(records);

                if (daLuu)
                {
                    MessageBox.Show(
                        "Đã lưu số Căn cước công dân.",
                        "Thành công",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    HienThiTree(_allRecords);
                }
            };

            panel.Controls.Add(lbl);
            panel.Controls.Add(txt);
            panel.Controls.Add(btn);
            return panel;
        }
        private void DisposeImages(Control control)
        {
            try
            {
                foreach (Control child
                         in control.Controls)
                {
                    if (child is PictureBox picture)
                    {
                        Image? image =
                            picture.Image;

                        picture.Image = null;

                        image?.Dispose();
                    }

                    DisposeImages(child);
                }
            }
            catch
            {
            }
        }

        private void SetThumbnail(PdfCardInfo card, byte[] imageBytes)
        {
            void SetImage()
            {
                try
                {
                    if (card.Panel.IsDisposed ||
                        card.PictureBox.IsDisposed)
                        return;

                    Image? image =
                        BytesToImage(imageBytes);

                    if (image == null)
                    {
                        card.StatusLabel.Text =
                            "Không đọc được";

                        return;
                    }

                    card.PictureBox.Controls.Clear();

                    Image? oldImage =
                        card.PictureBox.Image;

                    card.PictureBox.Image =
                        image;

                    oldImage?.Dispose();

                    card.Loaded = true;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                }
            }

            if (InvokeRequired)
                BeginInvoke(SetImage);
            else
                SetImage();
        }

        private async void flowPdf_Scroll(object? sender, ScrollEventArgs e)
        {
            if (_pdfLoadCts == null)
                return;

            try
            {
                await LoadVisibleThumbnailsAsync(
                    _pdfLoadCts.Token);

                // Nếu đã scroll gần cuối
                if (IsNearBottom())
                {
                    await LoadLastCardsAsync(
                        _pdfLoadCts.Token);
                }
            }
            catch (OperationCanceledException)
            {
            }
        }
        private async Task LoadLastCardsAsync(CancellationToken token)
        {
            List<PdfCardInfo> cards =
                _pdfCards.Values
                .Where(x => !x.Loaded)
                .TakeLast(10)
                .ToList();

            if (cards.Count == 0)
                return;

            List<Task> tasks = new();

            foreach (PdfCardInfo card in cards)
            {
                token.ThrowIfCancellationRequested();

                tasks.Add(
                    LoadThumbnailAsync(
                        card,
                        token));
            }

            await Task.WhenAll(tasks);
        }
        private bool IsNearBottom()
        {
            int verticalValue =
                flowPdf.VerticalScroll.Value;

            int clientHeight =
                flowPdf.ClientSize.Height;

            int maxValue =
                flowPdf.VerticalScroll.Maximum;

            return verticalValue + clientHeight
                   >= maxValue - 300;
        }

        private async Task LoadThumbnailAsync(PdfCardInfo card, CancellationToken token)
        {
            if (card.Loaded)
                return;

            if (!File.Exists(card.Path))
            {
                SetCardError(
                    card,
                    "Không tìm thấy PDF");

                return;
            }

            // ==========================================
            // TRÁNH 2 TASK CÙNG LOAD 1 PDF
            // ==========================================

            if (!_loadingThumbnails.TryAdd(
                card.Path,
                0))
            {
                return;
            }

            try
            {
                SetCardStatus(
                    card,
                    "Đang tải...");

                // ==========================================
                // CACHE DISK
                // ==========================================

                byte[]? cached =
                    await Task.Run(
                        () => LoadThumbnailCache(
                            card.Path),
                        token);

                token.ThrowIfCancellationRequested();

                if (cached != null &&
                    cached.Length > 0)
                {
                    SetThumbnail(
                        card,
                        cached);

                    return;
                }

                // ==========================================
                // GIỚI HẠN 4 PDF ĐỒNG THỜI
                // ==========================================

                await _pdfRenderSemaphore.WaitAsync(
                    token);

                try
                {
                    token.ThrowIfCancellationRequested();

                    byte[]? imageBytes =
                        await Task.Run(
                            () => RenderFirstPage(
                                card.Path),
                            token);

                    token.ThrowIfCancellationRequested();

                    if (imageBytes == null)
                    {
                        SetCardError(
                            card,
                            "Không render được");

                        card.Failed = true;

                        return;
                    }

                    // ======================================
                    // LƯU CACHE
                    // ======================================

                    await Task.Run(
                        () => SaveThumbnailCache(
                            card.Path,
                            imageBytes),
                        token);

                    // ======================================
                    // HIỂN THỊ
                    // ======================================

                    SetThumbnail(
                        card,
                        imageBytes);
                }
                finally
                {
                    _pdfRenderSemaphore.Release();
                }
            }
            catch (OperationCanceledException)
            {
            }
            catch (Exception ex)
            {
                Debug.WriteLine(
                    $"Thumbnail error: {card.Path}\n{ex}");

                SetCardError(
                    card,
                    "Lỗi");
            }
            finally
            {
                _loadingThumbnails.TryRemove(
                    card.Path,
                    out _);
            }
        }

        private void SetCardStatus(PdfCardInfo card, string text)
        {
            void Set()
            {
                if (card.Panel.IsDisposed)
                    return;

                card.StatusLabel.Text =
                    text;
            }

            if (InvokeRequired)
                BeginInvoke(Set);
            else
                Set();
        }

        private void SetCardError(PdfCardInfo card, string text)
        {
            void Set()
            {
                if (card.Panel.IsDisposed)
                    return;

                card.StatusLabel.Text =
                    text;

                card.StatusLabel.ForeColor =
                    System.Drawing.Color.Red;
            }

            if (InvokeRequired)
                BeginInvoke(Set);
            else
                Set();
        }

        private bool IsCardVisible(Control control)
        {
            if (!control.Visible)
                return false;

            Rectangle rect =
                flowPdf.ClientRectangle;

            rect.Inflate(
                300,
                500);

            return rect.IntersectsWith(
                control.Bounds);
        }

        private async Task LoadVisibleThumbnailsAsync(CancellationToken token)
        {
            if (_loadingPdfList)
                return;

            _loadingPdfList = true;

            try
            {
                List<PdfCardInfo> visibleCards =
                    new List<PdfCardInfo>();

                foreach (Control control in flowPdf.Controls)
                {
                    if (control is not Panel panel)
                        continue;

                    PdfCardInfo? card =
                        _pdfCards.Values
                        .FirstOrDefault(
                            x => ReferenceEquals(
                                x.Panel,
                                panel));

                    if (card == null)
                        continue;

                    if (card.Loaded)
                        continue;

                    if (!IsCardVisible(panel))
                        continue;

                    visibleCards.Add(card);
                }

                // ==========================================
                // LOAD SONG SONG
                // ==========================================

                List<Task> tasks =
                    new List<Task>();

                foreach (PdfCardInfo card in visibleCards)
                {
                    token.ThrowIfCancellationRequested();

                    tasks.Add(
                        LoadThumbnailAsync(
                            card,
                            token));
                }

                await Task.WhenAll(tasks);
            }
            catch (OperationCanceledException)
            {
            }
            finally
            {
                _loadingPdfList = false;
            }
        }

        private PdfCardInfo TaoPdfItem(Dictionary<string, object?> record, string path, string fileName)
        {
            Panel panel =
                new Panel
                {
                    Width = 350,
                    Height = 400,
                    Margin = new Padding(10),
                    Padding = new Padding(5),
                    BorderStyle = BorderStyle.FixedSingle,
                    BackColor = System.Drawing.Color.White,
                    Cursor = Cursors.Hand
                };

            PictureBox picturePdf =
                new PictureBox
                {
                    Dock = DockStyle.Top,
                    Height = 350,
                    SizeMode = PictureBoxSizeMode.Zoom,
                    BackColor = System.Drawing.Color.WhiteSmoke,
                    BorderStyle = BorderStyle.FixedSingle,
                    Cursor = Cursors.Hand
                };

            Label statusLabel =
                new Label
                {
                    Text = "Đang chờ...",
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.MiddleCenter,
                    ForeColor = System.Drawing.Color.Gray,
                    BackColor = System.Drawing.Color.WhiteSmoke
                };

            picturePdf.Controls.Add(
                statusLabel);

            Label lblFileName =
                new Label
                {
                    Text = fileName,
                    Dock = DockStyle.Bottom,
                    Height = 32,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = new Font(
                        "Segoe UI",
                        9F,
                        FontStyle.Bold),
                    AutoEllipsis = true,
                    Padding = new Padding(3),
                    Cursor = Cursors.Hand
                };

            // ==========================================
            // CLICK MỞ PDF
            // ==========================================

            panel.Click +=
                (s, e) => MoPdf(path);

            picturePdf.Click +=
                (s, e) => MoPdf(path);

            //lblFileName.Click +=
            //    (s, e) => MoPdf(path);
            // Double click tên file: đổi tên
            lblFileName.Click += (s, e) =>
            {
                DoiTenPdf(
                    record,
                    path,
                    lblFileName);
            };

            // ==========================================
            // ADD CONTROL
            // ==========================================

            panel.Controls.Add(
                lblFileName);

            panel.Controls.Add(
                picturePdf);

            return new PdfCardInfo
            {
                Panel = panel,
                PictureBox = picturePdf,
                FileNameLabel = lblFileName,
                StatusLabel = statusLabel,
                Path = path,
                FileName = fileName,
                Record = record
            };
        }

        private void DoiTenPdf(Dictionary<string, object?> record, string oldPath, Label label)
        {
            try
            {
                if (!File.Exists(oldPath))
                {
                    MessageBox.Show(
                        "Không tìm thấy file PDF:\n\n" + oldPath,
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }

                // ==========================================================
                // LẤY TÊN HIỆN TẠI TỪ JSON
                // KHÔNG LẤY TỪ oldPath
                // ==========================================================

                string currentFileName = GetValue(record, "TenFile");

                // Nếu JSON chưa có TenFile thì lấy từ file gốc
                if (string.IsNullOrWhiteSpace(currentFileName))
                {
                    currentFileName = Path.GetFileName(oldPath);
                }

                string oldNameWithoutExtension =Path.GetFileNameWithoutExtension(currentFileName);

                string extension = Path.GetExtension(oldPath);

                if (string.IsNullOrWhiteSpace(extension))
                    extension = ".pdf";

                // ==========================================================
                // NHẬP TÊN MỚI
                // ==========================================================
                List<FrmChonTenFile.DanhMucTaiLieu> danhSachTenFile = LayDanhSachTenFile();

                if (danhSachTenFile.Count == 0)
                {
                    MessageBox.Show(
                        "Không có danh sách tên file trong danh mục tài liệu.",
                        "Thông báo",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }
                using FrmChonTenFile frm =new FrmChonTenFile(currentFileName, danhSachTenFile);
                if (frm.ShowDialog(this) != DialogResult.OK)return;
                string newName = frm.TenDuocChon.Trim();
                if (string.IsNullOrWhiteSpace(newName))
                    return;
                // ==========================================================
                // NẾU NGƯỜI DÙNG NHẬP .PDF THÌ GIỮ NGUYÊN
                // ==========================================================
                if (!newName.EndsWith(
                    ".pdf",
                    StringComparison.OrdinalIgnoreCase))
                {
                    newName += extension;
                }

                // ==========================================================
                // KIỂM TRA TÊN FILE
                // ==========================================================

                foreach (char c in Path.GetInvalidFileNameChars())
                {
                    if (newName.Contains(c))
                    {
                        MessageBox.Show(
                            $"Tên file không hợp lệ.\n\n" +
                            $"Ký tự '{c}' không được phép.",
                            "Lỗi",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);

                        return;
                    }
                }

                // ==========================================================
                // KHÔNG ĐỔI NẾU TÊN GIỐNG TÊN ĐANG CÓ TRONG JSON
                // ==========================================================

                if (string.Equals(
                    currentFileName,
                    newName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                // ==========================================================
                // LẤY ID
                // ==========================================================

                string id = GetValue(
                    record,
                    "id");

                if (string.IsNullOrWhiteSpace(id))
                {
                    MessageBox.Show(
                        "Bản ghi không có ID.",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }

                // ==========================================================
                // LƯU JSON THEO ID
                // ==========================================================

                SaveMetadataJson(
                    record,
                    newName);

                // ==========================================================
                // CẬP NHẬT RECORD TRÊN RAM
                // ==========================================================

                record["TenFile"] = newName;

                // ==========================================================
                // CẬP NHẬT UI
                // ==========================================================

                label.Text = newName;

                MessageBox.Show(
                    "Đã cập nhật tên mới vào dữ liệu.\n\n" +
                    $"ID: {id}\n" +
                    $"Tên mới: {newName}\n\n" +
                    "File PDF gốc CHƯA bị đổi tên.",
                    "Đã lưu",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Không thể cập nhật tên PDF:\n\n" +
                    ex.Message,
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private byte[]? RenderFirstPage(string pdfPath)
        {
            try
            {
                if (!File.Exists(pdfPath))
                    return null;

                using FileStream stream = new FileStream(
                    pdfPath,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.ReadWrite);

                // Render trang đầu PDF
                using SKBitmap skBitmap = Conversion.ToImage(
                    stream,
                    page: 0);

                if (skBitmap == null)
                    return null;

                // ==========================================
                // GIỚI HẠN KÍCH THƯỚC THUMBNAIL
                // ==========================================

                const int maxWidth = 800;
                const int maxHeight = 1200;

                int width = skBitmap.Width;
                int height = skBitmap.Height;

                double scale = Math.Min(
                    (double)maxWidth / width,
                    (double)maxHeight / height);

                // Không phóng to ảnh nhỏ
                if (scale > 1)
                    scale = 1;

                int newWidth = Math.Max(
                    1,
                    (int)(width * scale));

                int newHeight = Math.Max(
                    1,
                    (int)(height * scale));

                SKBitmap finalBitmap;

                // Nếu đã đủ nhỏ thì dùng luôn
                if (newWidth == width &&
                    newHeight == height)
                {
                    finalBitmap = skBitmap.Copy();
                }
                else
                {
                    finalBitmap = skBitmap.Resize(
                        new SKImageInfo(
                            newWidth,
                            newHeight),
                        SKSamplingOptions.Default);

                    if (finalBitmap == null)
                        return null;
                }

                using (finalBitmap)
                {
                    using SKImage skImage =
                        SKImage.FromBitmap(finalBitmap);

                    using SKData skData =
                        skImage.Encode(
                            SKEncodedImageFormat.Jpeg,
                            75);

                    return skData.ToArray();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(
                    $"Render PDF lỗi: {pdfPath}\n{ex}");

                return null;
            }
        }
        private Image? BytesToImage(byte[] bytes)
        {
            try
            {
                using MemoryStream ms =
                    new MemoryStream(bytes);

                using Image temp =
                    Image.FromStream(ms);

                return new Bitmap(temp);
            }
            catch
            {
                return null;
            }
        }
        private void MoPdf(string path)
        {
            if (!File.Exists(path))
            {
                MessageBox.Show(
                    "Không tìm thấy file:\n\n" + path,
                    "File không tồn tại",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            try
            {
                Process.Start(
                    new ProcessStartInfo
                    {
                        FileName = path,
                        UseShellExecute = true
                    });
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Không thể mở PDF:\n\n" + ex.Message,
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        private string GetThumbnailDirectory()
        {
            string directory =
                Path.Combine(
                    AppContext.BaseDirectory,
                    "data",
                    "thumbnails");

            Directory.CreateDirectory(directory);

            return directory;
        }
        private string GetThumbnailPath(string pdfPath)
        {
            try
            {
                FileInfo info =
                    new FileInfo(pdfPath);

                string key =
                    pdfPath.ToLowerInvariant()
                    + "|"
                    + info.Length
                    + "|"
                    + info.LastWriteTimeUtc.Ticks;

                using SHA256 sha =
                    SHA256.Create();

                byte[] hash =
                    sha.ComputeHash(
                        Encoding.UTF8.GetBytes(key));

                string fileName =
                    Convert.ToHexString(hash) + ".jpg";

                return Path.Combine(
                    GetThumbnailDirectory(),
                    fileName);
            }
            catch
            {
                using SHA256 sha =
                    SHA256.Create();

                byte[] hash =
                    sha.ComputeHash(
                        Encoding.UTF8.GetBytes(
                            pdfPath.ToLowerInvariant()));

                return Path.Combine(
                    GetThumbnailDirectory(),
                    Convert.ToHexString(hash) + ".jpg");
            }
        }
        private byte[]? LoadThumbnailCache(string pdfPath)
        {
            try
            {
                string cachePath =
                    GetThumbnailPath(pdfPath);

                if (!File.Exists(cachePath))
                    return null;

                return File.ReadAllBytes(cachePath);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(
                    $"Đọc cache lỗi: {pdfPath}\n{ex}");

                return null;
            }
        }
        private void SaveThumbnailCache(string pdfPath, byte[] imageBytes)
        {
            try
            {
                string cachePath =
                    GetThumbnailPath(pdfPath);

                string tempPath =
                    cachePath + ".tmp";

                File.WriteAllBytes(
                    tempPath,
                    imageBytes);

                if (File.Exists(cachePath))
                    File.Delete(cachePath);

                File.Move(
                    tempPath,
                    cachePath);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(
                    $"Lưu cache lỗi: {pdfPath}\n{ex}");
            }
        }
        private void SaveMetadataJson(Dictionary<string, object?> record, string newFileName)
        {
            try
            {
                string dataFolder = Path.Combine(AppContext.BaseDirectory, "data");
                string jsonPath = Path.Combine(dataFolder, "metadata.json");
                if (!File.Exists(jsonPath))
                {
                    MessageBox.Show(
                        "Không tìm thấy metadata.json.",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }

                // ==========================================================
                // LẤY ID CỦA RECORD HIỆN TẠI
                // ==========================================================

                string id = GetValue(record, "id");

                if (string.IsNullOrWhiteSpace(id))
                {
                    MessageBox.Show(
                        "Bản ghi không có ID.",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }

                // ==========================================================
                // ĐỌC JSON
                // ==========================================================

                string json = File.ReadAllText(
                    jsonPath,
                    Encoding.UTF8);

                var data =
                    JsonConvert.DeserializeObject<
                        List<Dictionary<string, object?>>>(json)
                    ?? new List<Dictionary<string, object?>>();

                // ==========================================================
                // TÌM ĐÚNG BẢN GHI THEO ID
                // ==========================================================

                Dictionary<string, object?>? targetRecord =
                    data.FirstOrDefault(x =>
                        string.Equals(
                            GetValue(x, "id"),
                            id,
                            StringComparison.OrdinalIgnoreCase));

                if (targetRecord == null)
                {
                    MessageBox.Show(
                        $"Không tìm thấy bản ghi có ID: {id}",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }

                // ==========================================================
                // CHỈ SỬA TenFile
                // ==========================================================

                targetRecord["TenFile"] = newFileName;

                // ==========================================================
                // GHI LẠI JSON
                // ==========================================================

                string newJson =
                    JsonConvert.SerializeObject(
                        data,
                        Formatting.Indented);

                File.WriteAllText(
                    jsonPath,
                    newJson,
                    new UTF8Encoding(false));

                // ==========================================================
                // CẬP NHẬT RECORD ĐANG HIỂN THỊ
                // ==========================================================

                record["TenFile"] = newFileName;

                Debug.WriteLine(
                    $"Đã cập nhật TenFile cho ID={id}: {newFileName}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Không thể lưu metadata.json:\n\n" +
                    ex.Message,
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        private void TimKiemDangVien(string keyword)
        {
            keyword = keyword.Trim();

            // Không nhập gì -> hiện toàn bộ
            if (string.IsNullOrWhiteSpace(keyword))
            {
                HienThiTree(_allRecords);
                return;
            }

            // ==========================================
            // TÌM TRỰC TIẾP TRONG RAM
            // ==========================================

            var result = _allRecords
                .Where(x =>
                    GetValue(x, "DangVienNamSinh")
                        .Contains(
                            keyword,
                            StringComparison.OrdinalIgnoreCase))
                .ToList();

            Debug.WriteLine(
                $"Tìm [{keyword}] -> {result.Count} bản ghi");

            HienThiTree(result);
        }
        private void HienThiTree(List<Dictionary<string, object?>> records)
        {
            treeView1.BeginUpdate();

            try
            {
                treeView1.Nodes.Clear();

                // ==========================================================
                // CHỈ LẤY BẢN GHI CÓ TÊN CHI BỘ
                // ==========================================================

                var chiBoGroups = records
                    .Where(x =>
                        !string.IsNullOrWhiteSpace(
                            GetValue(x, "TenChiBo")))
                    .GroupBy(x =>
                        GetValue(x, "TenChiBo"))
                    .OrderBy(x => x.Key);

                foreach (var chiBoGroup in chiBoGroups)
                {
                    string tenChiBo = chiBoGroup.Key;

                    // ======================================================
                    // CHỈ LẤY ĐẢNG VIÊN CÓ TÊN
                    // ======================================================

                    var dangVienGroups = chiBoGroup
                        .Where(x =>
                            !string.IsNullOrWhiteSpace(
                                GetValue(x, "DangVienNamSinh")))
                        .GroupBy(x =>
                            GetValue(x, "DangVienNamSinh"))
                        .OrderBy(x => x.Key)
                        .ToList();

                    if (dangVienGroups.Count == 0)
                        continue;

                    // ======================================================
                    // CẤP 1: CHI BỘ
                    // ======================================================

                    TreeNode chiBoNode =
                        new TreeNode(
                            $"📁 {tenChiBo} ({dangVienGroups.Count})");

                    // Lưu toàn bộ record của chi bộ
                    chiBoNode.Tag =
                        dangVienGroups
                            .SelectMany(x => x)
                            .ToList();

                    // ======================================================
                    // CẤP 2: ĐẢNG VIÊN
                    // ======================================================

                    foreach (var dangVienGroup in dangVienGroups)
                    {
                        string dangVien = dangVienGroup.Key;

                        // --------------------------------------------------
                        // Kiểm tra CCCD
                        // --------------------------------------------------

                        var recordsDangVien = dangVienGroup.ToList();

                        bool daCoCCCD = recordsDangVien.All(x =>
                            !string.IsNullOrWhiteSpace(
                                GetValue(x, "CCCD")));

                        // --------------------------------------------------
                        // Tạo tên node
                        // --------------------------------------------------

                        string tenNode = daCoCCCD
                            ? $"👤 {dangVien}  ✔ Hoàn thành"
                            : $"👤 {dangVien}";

                        TreeNode dangVienNode =
                            new TreeNode(tenNode);

                        // Lưu toàn bộ PDF của đảng viên
                        dangVienNode.Tag = recordsDangVien;

                        // Có CCCD thì đổi màu để dễ nhìn
                        if (daCoCCCD)
                        {
                            dangVienNode.ForeColor = System.Drawing.Color.Green;
                        }

                        chiBoNode.Nodes.Add(dangVienNode);
                    }

                    treeView1.Nodes.Add(chiBoNode);
                }

                // ==========================================================
                // MỞ TOÀN BỘ CHI BỘ
                // ==========================================================

                foreach (TreeNode node in treeView1.Nodes)
                {
                    node.Expand();
                }
            }
            finally
            {
                treeView1.EndUpdate();
            }
        }
        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            TimKiemDangVien(txtSearch.Text);
        }
        private void btnChonThuMuc_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Chọn thư mục";
                dialog.ShowNewFolderButton = true;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedFolder = dialog.SelectedPath;

                    // Ví dụ: hiển thị đường dẫn lên TextBox
                    txtpaththumuc.Text = selectedFolder;
                }
            }
        }
        private void btnChonDanhMuc_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "Chọn file danh mục tài liệu";
                dialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls";
                dialog.Multiselect = false;

                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                string excelPath = dialog.FileName;

                try
                {
                    using (var workbook = new XLWorkbook(excelPath))
                    {
                        var worksheet = workbook.Worksheet(1);

                        var rows = worksheet.RowsUsed().ToList();

                        if (rows.Count < 2)
                        {
                            MessageBox.Show(
                                "File Excel không có dữ liệu.",
                                "Thông báo",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);

                            return;
                        }

                        // ==========================================
                        // LẤY HEADER VÀ BỎ DẤU TIẾNG VIỆT
                        // ==========================================

                        var headerRow = rows[0];

                        List<string> headers = new List<string>();

                        foreach (var cell in headerRow.CellsUsed())
                        {
                            string header = cell.GetString().Trim();

                            // Bỏ dấu tiếng Việt
                            header = Utils.NormalizeKey(header);

                            headers.Add(header);
                        }

                        if (headers.Count == 0)
                        {
                            MessageBox.Show(
                                "Không tìm thấy tiêu đề cột trong file Excel.",
                                "Lỗi",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);

                            return;
                        }

                        // ==========================================
                        // ĐỌC DANH SÁCH BẢN GHI
                        // ==========================================

                        List<Dictionary<string, object?>> records =
                            new List<Dictionary<string, object?>>();

                        for (int i = 1; i < rows.Count; i++)
                        {
                            var row = rows[i];

                            // Bỏ qua dòng trống
                            if (row.CellsUsed().All(c =>
                                string.IsNullOrWhiteSpace(c.GetString())))
                            {
                                continue;
                            }

                            Dictionary<string, object?> record =
                                new Dictionary<string, object?>();

                            for (int j = 0; j < headers.Count; j++)
                            {
                                string header = headers[j];

                                if (string.IsNullOrWhiteSpace(header))
                                    continue;

                                var cell = row.Cell(j + 1);

                                string value = cell.GetString().Trim();

                                record[header] =
                                    string.IsNullOrWhiteSpace(value)
                                        ? null
                                        : value;
                            }

                            records.Add(record);
                        }

                        // ==========================================
                        // TẠO THƯ MỤC DATA
                        // ==========================================

                        string? directory =
                            Path.GetDirectoryName(danhMucJsonPath);

                        if (!string.IsNullOrEmpty(directory))
                        {
                            Directory.CreateDirectory(directory);
                        }

                        // ==========================================
                        // LƯU JSON
                        // ==========================================

                        var options = new JsonSerializerOptions
                        {
                            WriteIndented = true,

                            // Giữ nguyên tiếng Việt trong VALUE
                            Encoder = System.Text.Encodings.Web
                                .JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                        };

                        string json =
                            System.Text.Json.JsonSerializer.Serialize(
                                records,
                                options);

                        File.WriteAllText(
                            danhMucJsonPath,
                            json,
                            Encoding.UTF8);

                        danhMucTaiLieu = records;
                        HienThiDanhMuc(danhMucTaiLieu);
                        txtpathdanhmuc.Text = danhMucJsonPath;
                        MessageBox.Show(
                            $"Import danh mục thành công!\n\n" +
                            $"Số bản ghi: {records.Count:N0}\n\n" +
                            $"Đã lưu tại:\n{danhMucJsonPath}",
                            "Import thành công",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        $"Không thể import danh mục.\n\n" +
                        $"Chi tiết:\n{ex.Message}",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }
        private void LoadDanhMucVaoDataGridView()
        {
            try
            {
                if (!File.Exists(danhMucJsonPath))
                {
                    return;
                }

                string json = File.ReadAllText(
                    danhMucJsonPath,
                    Encoding.UTF8);

                danhMucTaiLieu = System.Text.Json.JsonSerializer.Deserialize<List<Dictionary<string, object?>>>(json) ?? new List<Dictionary<string, object?>>();

                HienThiDanhMuc(danhMucTaiLieu);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Không thể load danh mục:\n{ex.Message}",
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }

        }
        private void HienThiDanhMuc(List<Dictionary<string, object?>> records)
        {
            dataGridView.DataSource = null;
            dataGridView.Columns.Clear();

            if (records.Count == 0)
                return;

            // Lấy toàn bộ key
            var keys = records
                .SelectMany(x => x.Keys)
                .Distinct()
                .ToList();

            // Tạo cột
            foreach (string key in keys)
            {
                dataGridView.Columns.Add(
                    key,
                    key);
            }

            // Thêm dữ liệu
            foreach (var record in records)
            {
                int rowIndex =
                    dataGridView.Rows.Add();

                foreach (string key in keys)
                {
                    if (record.TryGetValue(key, out var value))
                    {
                        dataGridView.Rows[rowIndex]
                            .Cells[key]
                            .Value = value?.ToString() ?? "";
                    }
                }
            }

            dataGridView.AutoSizeColumnsMode =
                DataGridViewAutoSizeColumnsMode.None;

            dataGridView.AllowUserToAddRows = false;
            dataGridView.ReadOnly = true;
        }
        private void btnSaveConfig_Click(object sender, EventArgs e)
        {
            LuuCauHinh();
        }
        private void LuuCauHinh()
        {
            try
            {
                var settings = new AppSettings
                {
                    ThuMuc = txtpaththumuc.Text.Trim(),
                    DanhMuc = txtpathdanhmuc.Text.Trim(),
                    Url = txturlapi.Text.Trim()
                };

                string? directory =
                    Path.GetDirectoryName(settingsJsonPath);

                if (!string.IsNullOrEmpty(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web
                        .JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };

                string json = System.Text.Json.JsonSerializer.Serialize(settings, options);

                File.WriteAllText(
                    settingsJsonPath,
                    json,
                    Encoding.UTF8);

                MessageBox.Show(
                   $"Lưu thành công",
                   "Thành công",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Không thể lưu cấu hình:\n{ex.Message}",
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        private void LoadCauHinh()
        {
            try
            {
                if (!File.Exists(settingsJsonPath))
                    return;

                string json =
                    File.ReadAllText(
                        settingsJsonPath,
                        Encoding.UTF8);

                var settings = System.Text.Json.JsonSerializer.Deserialize<AppSettings>(json);

                if (settings == null)
                    return;

                txtpaththumuc.Text = settings.ThuMuc;
                txtpathdanhmuc.Text = settings.DanhMuc;
                txturlapi.Text = settings.Url;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Không thể load cấu hình:\n{ex.Message}",
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        private async void btnNhanDang_Click(object sender, EventArgs e)
        {
            if (selectedDangVienRecords.Count == 0)
            {
                MessageBox.Show(
                    "Vui lòng chọn một đảng viên trước.",
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }
            btnNhanDang.Enabled = false;
            try
            {
                int total = selectedDangVienRecords.Count;
                int success = 0;

                for (int i = 0; i < total; i++)
                {
                    var record = selectedDangVienRecords[i];

                    Text =
                        $"Đang nhận dạng {i + 1}/{total}...";

                    bool result = await NhanDangMotRecordAsync(record);

                    if (result)
                        success++;
                }

                HienThiDanhSachPdf(
                    selectedDangVienRecords);

                MessageBox.Show(
                    $"Đã nhận dạng xong.\n\n" +
                    $"Tổng: {total}\n" +
                    $"Thành công: {success}\n" +
                    $"Không xác định: {total - success}",
                    "Hoàn thành",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            finally
            {
                btnNhanDang.Enabled = false;
                Text = "renamePDF_v2";
            }

        }
        private async Task<bool> NhanDangMotRecordAsync(Dictionary<string, object?> record)
        {
            // Bước 1
            string nhomVanBan = GetNhomVanBan(record);

            if (string.IsNullOrWhiteSpace(nhomVanBan))
                return false;

            // Bước 2
            var danhMucLoc = LocDanhMucTheoNhom(nhomVanBan);

            if (danhMucLoc.Count == 0)
                return false;
            string PhanLoai = GetValue(record, "phan_loai");
            if (string.IsNullOrWhiteSpace(PhanLoai) && PhanLoai == "BIA")
                return false;

            // Bước 3
            string? tenfile = await GuiDeepSeekNhanDangAsync(record, danhMucLoc);

            if (string.IsNullOrWhiteSpace(tenfile))
                return false;
            record["TenFile"] = tenfile;
            string id = GetValue(record, "id");
            if (string.IsNullOrWhiteSpace(id))
            {
                Debug.WriteLine(
                    "Không tìm thấy id của bản ghi.");

                return false;
            }
            bool updated = await UpdateRecordToJsonAsync(id, tenfile);
            if (!updated)
            {
                Debug.WriteLine(
                    $"Không cập nhật được JSON. ID: {id}");

                return false;
            }
            return true;
        }
        private async Task<bool> UpdateRecordToJsonAsync(string id, string tenFile)
        {
            try
            {
                string jsonPath = metadataJsonPath;

                if (!File.Exists(jsonPath))
                {
                    Debug.WriteLine(
                        $"Không tìm thấy file JSON: {jsonPath}");

                    return false;
                }

                // =====================================================
                // ĐỌC JSON
                // =====================================================

                string json = await File.ReadAllTextAsync(jsonPath, Encoding.UTF8);
                var records = System.Text.Json.JsonSerializer.Deserialize<List<Dictionary<string, object?>>>(json);
                if (records == null)
                    return false;

                // =====================================================
                // TÌM BẢN GHI THEO ID
                // =====================================================

                var record = records.FirstOrDefault(x => GetValue(x, "id").Equals(id, StringComparison.OrdinalIgnoreCase));

                if (record == null)
                {
                    Debug.WriteLine(
                        $"Không tìm thấy bản ghi ID = {id}");

                    return false;
                }
                // =====================================================
                // UPDATE TENFILE
                // =====================================================
                record["TenFile"] = tenFile;
                // =====================================================
                // GHI LẠI JSON
                // =====================================================

                var options =
                    new JsonSerializerOptions
                    {
                        WriteIndented = true,

                        Encoder =
                            System.Text.Encodings.Web
                                .JavaScriptEncoder
                                .UnsafeRelaxedJsonEscaping
                    };

                string newJson = System.Text.Json.JsonSerializer.Serialize(
                        records,
                        options);

                await File.WriteAllTextAsync(
                    jsonPath,
                    newJson,
                    Encoding.UTF8);

                Debug.WriteLine(
                    $"Đã cập nhật ID={id} -> TenFile={tenFile}");

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(
                    $"Lỗi UpdateRecordToJsonAsync: {ex}");

                return false;
            }
        }
        private string GetNhomVanBan(Dictionary<string, object?> record)
        {
            if (record.TryGetValue("NhomVanBan", out object? value))
            {
                string nhomVanBan = value?.ToString()?.Trim() ?? "";

                if (string.IsNullOrWhiteSpace(nhomVanBan))
                    return "";

                string so = nhomVanBan.Split('_')[0];

                return int.TryParse(so, out int number)
                    ? number.ToString()
                    : so;
            }

            return "";
        }
        private List<Dictionary<string, object?>> LocDanhMucTheoNhom(string nhomVanBan)
        {
            return danhMucTaiLieu.ToList();
            //return danhMucTaiLieu.Where(x =>x.TryGetValue("do_uu_tien", out object? value)&& string.Equals(value?.ToString()?.Trim(),nhomVanBan,StringComparison.OrdinalIgnoreCase)).ToList();
        }
        private async Task<string?> GuiDeepSeekNhanDangAsync(Dictionary<string, object?> record, List<Dictionary<string, object?>> danhMucLoc)
        {
            if (_deepSeekClient == null)
                return null;

            if (danhMucLoc == null || danhMucLoc.Count == 0)
                return null;
            var thongTin = new Dictionary<string, string>
            {
                ["ten_loai_van_ban"] = GetValue(record, "ten_loai_van_ban"),

                ["so_thu_tu_van_ban"] = GetValue(record, "so_thu_tu_van_ban"),

                ["ngay_van_ban"] = GetValue(record, "ngay_van_ban"),

                ["tac_gia_van_ban"] = GetValue(record, "tac_gia_van_ban"),

                ["trich_yeu_noi_dung"] = GetValue(record, "trich_yeu_noi_dung"),

                ["but_tich"] = GetValue(record, "but_tich"),

                ["so_trang_van_ban"] = GetValue(record, "so_trang_van_ban")
            };
            danhMucLoc = danhMucLoc.ToList();
            var jsonOptions = new JsonSerializerOptions
            {
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            string thongTinJson = System.Text.Json.JsonSerializer.Serialize(thongTin, jsonOptions);

            string danhSachJson = System.Text.Json.JsonSerializer.Serialize(danhMucLoc, jsonOptions);
            string systemPrompt = """
Bạn là hệ thống phân loại tài liệu hồ sơ đảng viên.

NHIỆM VỤ:
Đối chiếu THÔNG TIN BẢN GHI với DANH MỤC TÀI LIỆU và chọn đúng 1 tài liệu phù hợp nhất.

DANH MỤC đã được lọc đúng nhóm văn bản.

NGUYÊN TẮC:

1. Tập trung vào TÊN LOẠI VĂN BẢN và nội dung cốt lõi.

2. Không cần giống nguyên văn.
Có thể bỏ qua:
- họ tên
- năm sinh
- năm tài liệu
- ngày tháng
- số hiệu
- nội dung mô tả thêm
- ghi chú
- "không có gì thay đổi"
- thông tin thu nhập, chuyên môn, nhiệm vụ...

3. Ví dụ:

"LÝ LỊCH ĐẢNG VIÊN _ VÌ VĂN DÍNH 1958"
→ "Lý lịch đảng viên"

"PHIẾU BỔ SUNG HỒ SƠ ĐẢNG VIÊN NĂM 2022 VÌ VĂN DÍNH 1958 KHÔNG CÓ GÌ THAY ĐỔI"
→ "Phiếu bổ sung hồ sơ đảng viên"

"BẢN KIỂM ĐIỂM CÁ NHÂN NĂM 2023 VÌ VĂN DÍNH 1958"
→ "Bản kiểm điểm cá nhân"

4. Ưu tiên theo thứ tự:
- ten_loai_van_ban
- trich_yeu_noi_dung
- so_thu_tu_van_ban
- ngay_van_ban
- tac_gia_van_ban
- but_tich
- so_trang_van_ban

5. Chỉ được chọn tài liệu có trong DANH MỤC.

6. "tenFile" phải lấy NGUYÊN VĂN từ trường "ten_file" của danh mục.

7. "tenTaiLieu" phải lấy từ chính bản ghi danh mục được chọn.

8. "tt" phải lấy từ chính bản ghi danh mục được chọn.

9. Luôn chọn ứng viên phù hợp nhất nếu có thể xác định.
Chỉ trả rỗng khi thực sự không có tài liệu nào phù hợp.

CHỈ TRẢ VỀ JSON.
""";

            string userPrompt = string.Format(@"
                    DANH MỤC TÀI LIỆU
                    ==================================================
                    {0}
                    ==================================================
                    THÔNG TIN BẢN GHI CẦN PHÂN LOẠI
                    ==================================================
                    {1}
                    ==================================================
                    Hãy chọn tài liệu phù hợp nhất.
                    Kết quả JSON:
                    {{
                        ""tt"": 0,
                        ""tenTaiLieu"": """",
                        ""tenFile"": """"
                    }}
                    ", danhSachJson, thongTinJson);
            // =========================================================
            // 5. GỌI DEEPSEEK
            // =========================================================

            try
            {
                string? response =
                    await _deepSeekClient.ChatAsync(
                        systemPrompt,
                        userPrompt);

                if (string.IsNullOrWhiteSpace(response))
                    return null;
                // =====================================================
                // 6. ĐỌC JSON
                // =====================================================
                using JsonDocument document = JsonDocument.Parse(response);
                if (!document.RootElement.TryGetProperty("tenFile", out JsonElement tenFileElement))
                {
                    return null;
                }
                string? tenFile = tenFileElement.GetString();
                if (string.IsNullOrWhiteSpace(tenFile))
                    return null;

                tenFile = tenFile.Trim();

                // =====================================================
                // 7. KIỂM TRA TEN FILE CÓ THỰC SỰ TRONG DANH MỤC
                // =====================================================

                bool tonTai =
                    danhMucLoc.Any(x =>
                        string.Equals(
                            GetValue(x, "ten_file"),
                            tenFile,
                            StringComparison.OrdinalIgnoreCase));

                if (!tonTai)
                {
                    Debug.WriteLine(
                        $"DeepSeek trả về tên file không tồn tại: {tenFile}");

                    return null;
                }

                // =====================================================
                // 8. TRẢ VỀ TÊN FILE
                // =====================================================

                return tenFile;
            }
            catch (System.Text.Json.JsonException ex)
            {
                Debug.WriteLine(
                    "DeepSeek trả JSON không hợp lệ:\n" +
                    ex.Message);

                return null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(
                    "GuiDeepSeekNhanDangAsync lỗi:\n" +
                    ex);

                return null;
            }
        }
        private async Task XuLyTrungTenFileAsync(ImportProgressForm progressForm)
        {
            // =====================================================
            // XÁC ĐỊNH DỮ LIỆU
            // =====================================================

            List<Dictionary<string, object?>> recordsToProcess;

            if (selectedDangVienRecords != null &&
                selectedDangVienRecords.Count > 0)
            {
                // Có chọn đảng viên
                recordsToProcess = selectedDangVienRecords;
            }
            else
            {
                // Không chọn -> xử lý toàn bộ
                recordsToProcess = _allRecords;
            }

            if (recordsToProcess == null || recordsToProcess.Count == 0)
            {
                MessageBox.Show(
                    "Không có dữ liệu để xử lý.",
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            // =====================================================
            // GROUP THEO ĐẢNG VIÊN
            // =====================================================

            var groups = recordsToProcess
                .GroupBy(x => new
                {
                    TenChiBo =
                        GetValue(x, "TenChiBo"),

                    DangVienNamSinh =
                        GetValue(x, "DangVienNamSinh")
                })
                .ToList();

            int soLuongDaDoi = 0;

            int tongSoBanGhi =
                recordsToProcess.Count;

            int daXuLy = 0;

            // =====================================================
            // HIỂN THỊ BAN ĐẦU
            // =====================================================

            progressForm.UpdateProgress(0, tongSoBanGhi, "Đang chuẩn bị...", "Đang xử lý");

            // =====================================================
            // XỬ LÝ TỪNG ĐẢNG VIÊN
            // =====================================================

            foreach (var group in groups)
            {
                Dictionary<string, int> counters =
                    new Dictionary<string, int>(
                        StringComparer.OrdinalIgnoreCase);

                foreach (var record in group)
                {
                    string tenFile =
                        GetValue(record, "TenFile").Trim();

                    if (!string.IsNullOrWhiteSpace(tenFile))
                    {
                        // File đầu tiên giữ nguyên
                        if (!counters.ContainsKey(tenFile))
                        {
                            counters[tenFile] = 0;
                        }
                        else
                        {
                            // File bị trùng
                            counters[tenFile]++;

                            int stt =
                                counters[tenFile];

                            string tenFileMoi = Utils.TaoTenFileTrung(tenFile,stt);

                            record["TenFile"] =
                                tenFileMoi;

                            soLuongDaDoi++;
                        }
                    }

                    // =================================================
                    // CẬP NHẬT TIẾN ĐỘ
                    // =================================================

                    daXuLy++;

                    progressForm.UpdateProgress(daXuLy, tongSoBanGhi, tenFile, "Đang xử lý");
                    // Cho UI cập nhật
                    await Task.Yield();
                }
            }

            // =====================================================
            // LƯU JSON
            // =====================================================

            progressForm.UpdateProgress(
                tongSoBanGhi,
                tongSoBanGhi,
                "Đang cập nhật JSON...");

            bool daLuuJson = await Task.Run(() =>
                    LuuTenFileVaoJson(
                        recordsToProcess));

            // =====================================================
            // CẬP NHẬT GIAO DIỆN
            // =====================================================

            progressForm.UpdateProgress(tongSoBanGhi, tongSoBanGhi, "Đang cập nhật giao diện...");

            if (selectedDangVienRecords != null &&
                selectedDangVienRecords.Count > 0)
            {
                HienThiDanhSachPdf(selectedDangVienRecords);
            }
            else
            {
                HienThiTree(_allRecords);
            }

            // =====================================================
            // HOÀN THÀNH
            // =====================================================
            progressForm.UpdateProgress(tongSoBanGhi, tongSoBanGhi, "Đã hoàn thành");
            // =====================================================
            // THÔNG BÁO
            // =====================================================

            MessageBox.Show(
                $"Đã xử lý xong.\n\n" +
                $"Số đảng viên: {groups.Count}\n" +
                $"Tổng bản ghi: {recordsToProcess.Count}\n" +
                $"Số tên file đã đổi: {soLuongDaDoi}\n" +
                $"Cập nhật JSON: {(daLuuJson ? "Thành công" : "Thất bại")}",
                "Hoàn thành",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        private bool LuuTenFileVaoJson(List<Dictionary<string, object?>> recordsCanCapNhat)
        {
            try
            {
                if (recordsCanCapNhat == null ||
                    recordsCanCapNhat.Count == 0)
                {
                    return false;
                }

                // =====================================================
                // KIỂM TRA FILE JSON
                // =====================================================

                if (string.IsNullOrWhiteSpace(metadataJsonPath))
                {
                    MessageBox.Show(
                        "Chưa xác định được đường dẫn file JSON.",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                    return false;
                }

                if (!File.Exists(metadataJsonPath))
                {
                    MessageBox.Show(
                        $"Không tìm thấy file JSON:\n\n{metadataJsonPath}",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                    return false;
                }

                // =====================================================
                // ĐỌC JSON
                // =====================================================

                string json =
                    File.ReadAllText(
                        metadataJsonPath,
                        Encoding.UTF8);

                var allRecords = System.Text.Json.JsonSerializer.Deserialize<List<Dictionary<string, object?>>>(json);

                if (allRecords == null ||
                    allRecords.Count == 0)
                {
                    return false;
                }

                // =====================================================
                // TẠO INDEX ID
                // Chỉ tìm mỗi ID 1 lần
                // =====================================================

                var recordsById =
                    new Dictionary<string,
                        Dictionary<string, object?>>(
                        StringComparer.OrdinalIgnoreCase);

                foreach (var record in allRecords)
                {
                    string id =
                        GetValue(record, "id").Trim();

                    if (string.IsNullOrWhiteSpace(id))
                        continue;

                    if (!recordsById.ContainsKey(id))
                    {
                        recordsById[id] = record;
                    }
                }

                // =====================================================
                // CẬP NHẬT
                // =====================================================

                int updatedCount = 0;

                foreach (var record in recordsCanCapNhat)
                {
                    string id =
                        GetValue(record, "id").Trim();

                    if (string.IsNullOrWhiteSpace(id))
                        continue;

                    if (!recordsById.TryGetValue(
                            id,
                            out var jsonRecord))
                    {
                        continue;
                    }

                    string tenFile =
                        GetValue(record, "TenFile").Trim();

                    jsonRecord["TenFile"] =
                        tenFile;

                    updatedCount++;
                }

                // =====================================================
                // GHI JSON
                // =====================================================

                var options =
                    new JsonSerializerOptions
                    {
                        WriteIndented = true,

                        Encoder =
                            System.Text.Encodings.Web
                                .JavaScriptEncoder
                                .UnsafeRelaxedJsonEscaping
                    };

                string newJson =
                    System.Text.Json.JsonSerializer.Serialize(
                        allRecords,
                        options);

                File.WriteAllText(
                    metadataJsonPath,
                    newJson,
                    Encoding.UTF8);

                Debug.WriteLine(
                    $"Đã cập nhật {updatedCount} bản ghi vào JSON.");

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Không thể cập nhật file JSON:\n\n" +
                    ex.Message,
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                return false;
            }
        }
        private bool LuuCapNhatRecordsVaoJson(List<Dictionary<string, object?>> recordsCanCapNhat)
        {
            try
            {
                if (recordsCanCapNhat == null ||
                    recordsCanCapNhat.Count == 0)
                {
                    return false;
                }

                if (string.IsNullOrWhiteSpace(metadataJsonPath))
                {
                    MessageBox.Show(
                        "Chưa xác định được đường dẫn file JSON.",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                    return false;
                }

                if (!File.Exists(metadataJsonPath))
                {
                    MessageBox.Show(
                        $"Không tìm thấy file JSON:\n\n{metadataJsonPath}",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                    return false;
                }

                // =====================================================
                // ĐỌC JSON
                // =====================================================

                string json = File.ReadAllText(
                    metadataJsonPath,
                    Encoding.UTF8);

                var allRecords = System.Text.Json.JsonSerializer.Deserialize<
                        List<Dictionary<string, object?>>>(json);

                if (allRecords == null)
                    return false;

                // =====================================================
                // TẠO INDEX THEO ID
                // =====================================================

                var recordsById =
                    new Dictionary<
                        string,
                        Dictionary<string, object?>>(
                            StringComparer.OrdinalIgnoreCase);

                foreach (var item in allRecords)
                {
                    string id =
                        GetValue(item, "id").Trim();

                    if (string.IsNullOrWhiteSpace(id))
                        continue;

                    if (!recordsById.ContainsKey(id))
                    {
                        recordsById[id] = item;
                    }
                }

                // =====================================================
                // CẬP NHẬT RECORD
                // =====================================================

                int updatedCount = 0;

                foreach (var record in recordsCanCapNhat)
                {
                    string id =
                        GetValue(record, "id").Trim();

                    if (string.IsNullOrWhiteSpace(id))
                        continue;

                    if (!recordsById.TryGetValue(
                            id,
                            out var jsonRecord))
                    {
                        continue;
                    }

                    // Cập nhật TOÀN BỘ các field
                    foreach (var field in record)
                    {
                        jsonRecord[field.Key] =
                            field.Value;
                    }

                    updatedCount++;
                }

                // =====================================================
                // GHI LẠI JSON
                // =====================================================

                var options =
                    new JsonSerializerOptions
                    {
                        WriteIndented = true,

                        Encoder =
                            System.Text.Encodings.Web
                                .JavaScriptEncoder
                                .UnsafeRelaxedJsonEscaping
                    };

                string newJson = System.Text.Json.JsonSerializer.Serialize(
                        allRecords,
                        options);

                File.WriteAllText(
                    metadataJsonPath,
                    newJson,
                    Encoding.UTF8);

                Debug.WriteLine(
                    $"Đã cập nhật {updatedCount} bản ghi.");

                return updatedCount > 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Không thể cập nhật JSON:\n\n" +
                    ex.Message,
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                return false;
            }
        }
       
        private async void btnXuLyDuLieu_Click(object sender, EventArgs e)
        {
            // =====================================================
            // XÁC NHẬN
            // =====================================================

            DialogResult result = MessageBox.Show(
                "Bạn có chắc chắn muốn xử lý trùng tên file không?\n\n" +
                "• Nếu đã chọn đảng viên: chỉ xử lý các đảng viên được chọn.\n" +
                "• Nếu chưa chọn đảng viên: xử lý toàn bộ đảng viên.\n\n" +
                "Các tên file trùng sẽ được đổi thành:\n" +
                "TênFile (1), TênFile (2), TênFile (3)...",
                "Xác nhận xử lý",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Question);

            if (result != DialogResult.OK)
                return;

            // =====================================================
            // KHÓA NÚT
            // =====================================================

            btnXuLyDuLieu.Enabled = false;

            ImportProgressForm? progressForm = null;

            try
            {
                // =================================================
                // TẠO FORM TIẾN ĐỘ
                // =================================================

                progressForm = new ImportProgressForm();

                progressForm.UpdateProgress(
                    0,
                    1,
                    "Đang chuẩn bị...");

                progressForm.Show(this);

                // =================================================
                // CHẠY XỬ LÝ
                // =================================================

                await XuLyTrungTenFileAsync(progressForm);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Có lỗi trong quá trình xử lý:\n\n" +
                    ex.Message,
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                // =================================================
                // LUÔN ĐÓNG FORM TIẾN ĐỘ
                // =================================================

                if (progressForm != null &&
                    !progressForm.IsDisposed)
                {
                    progressForm.Close();
                    progressForm.Dispose();
                }

                // =================================================
                // MỞ LẠI NÚT
                // =================================================

                btnXuLyDuLieu.Enabled = true;
            }
        }

        private async void btnSave_Click(object sender, EventArgs e)
        {
            // =====================================================
            // XÁC ĐỊNH DỮ LIỆU CẦN XỬ LÝ
            // =====================================================

            List<Dictionary<string, object?>> recordsToProcess;

            if (selectedDangVienRecords != null &&
                selectedDangVienRecords.Count > 0)
            {
                // Có chọn đảng viên
                recordsToProcess = selectedDangVienRecords;
            }
            else
            {
                // Không chọn -> xử lý toàn bộ
                recordsToProcess = _allRecords;
            }

            if (recordsToProcess == null ||
                recordsToProcess.Count == 0)
            {
                MessageBox.Show(
                    "Không có dữ liệu để xử lý.",
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            // =====================================================
            // XÁC NHẬN
            // =====================================================

            DialogResult result = MessageBox.Show(
                "Bạn có chắc chắn muốn lưu toàn bộ file không?\n\n" +
                "File sẽ được chuyển vào thư mục:\n" +
                "CCCD_Tên đảng viên\\TênFile",
                "Xác nhận lưu",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Question);

            if (result != DialogResult.OK)
                return;

            btnSave.Enabled = false;

            ImportProgressForm? progressForm = null;

            try
            {
                // =================================================
                // ĐỌC ĐƯỜNG DẪN TỪ SETTING.JSON
                // =================================================
                string rootPath = Utils.LayThuMucTuSetting();
                if (!Directory.Exists(rootPath))
                {
                    Directory.CreateDirectory(rootPath);
                }
                // =================================================
                // TẠO FORM TIẾN ĐỘ
                // =================================================
                progressForm = new ImportProgressForm();
                progressForm.Show(this);
                progressForm.UpdateProgress(0,recordsToProcess.Count,"Đang chuẩn bị...","Chuyển đổi file PDF");
                // =================================================
                // XỬ LÝ
                // =================================================
                int daXuLy = 0;
                int daChuyen = 0;
                int khongTonTai = 0;
                int biTrung = 0;
                int loi = 0;
                int total = recordsToProcess.Count;

                foreach (var record in recordsToProcess)
                {
                    daXuLy++;

                    string cccd = GetValue(record, "CCCD").Trim();
                    string tenDangVien = GetValue(record, "DangVienNamSinh").Trim();
                    tenDangVien = Utils.LayTenDangVien(tenDangVien);
                    string tenFile = GetValue(record, "TenFile").Trim();

                    // =================================================
                    // LẤY ĐƯỜNG DẪN FILE GỐC
                    // =================================================

                    string pathCu = GetValue(record, "duong_dan");
                    if (string.IsNullOrWhiteSpace(pathCu))
                    {
                        pathCu = GetValue(record, "FOLDER");
                    }
                    pathCu = pathCu.Trim();

                    // =================================================
                    // KIỂM TRA DỮ LIỆU
                    // =================================================

                    if (string.IsNullOrWhiteSpace(cccd) || string.IsNullOrWhiteSpace(tenDangVien) ||string.IsNullOrWhiteSpace(tenFile) || string.IsNullOrWhiteSpace(pathCu))
                    {
                        loi++;
                        progressForm.UpdateProgress(daXuLy,total,"Thiếu thông tin", "Chuyển đổi file PDF");
                        continue;
                    }

                    // =================================================
                    // LÀM SẠCH TÊN THƯ MỤC
                    // =================================================

                    string tenThuMuc = Utils.LamSachTenThuMuc($"{cccd}_{tenDangVien}");

                    // =================================================
                    // TẠO THƯ MỤC
                    // =================================================
                    string folderDangVien =Path.Combine(rootPath,tenThuMuc);
                    Directory.CreateDirectory(folderDangVien);
                    // =================================================
                    // ĐƯỜNG DẪN FILE MỚI
                    // =================================================
                    string pathMoi = Path.Combine(folderDangVien,tenFile);
                    // =================================================
                    // KIỂM TRA FILE GỐC
                    // =================================================
                    if (!File.Exists(pathCu))
                    {
                        khongTonTai++;
                        progressForm.UpdateProgress(daXuLy,total,$"Không tồn tại: {tenFile}", "Chuyển đổi file PDF");
                        continue;
                    }

                    // =================================================
                    // FILE ĐÍCH ĐÃ TỒN TẠI
                    // =================================================

                    if (File.Exists(pathMoi))
                    {
                        biTrung++;
                        progressForm.UpdateProgress(
                            daXuLy,
                            total,
                            $"Đã tồn tại: {tenFile}");

                        continue;
                    }

                    // =================================================
                    // COPY FILE
                    // =================================================

                    File.Copy(pathCu,pathMoi);

                    // =================================================
                    // CẬP NHẬT ĐƯỜNG DẪN MỚI
                    // =================================================
                    record["duong_dan"] =pathMoi;
                    daChuyen++;
                    // =================================================
                    // CẬP NHẬT TIẾN ĐỘ
                    // =================================================
                    progressForm.UpdateProgress(
                        daXuLy,
                        total,
                        tenFile, "Chuyển đổi file PDF");
                    if (daXuLy % 20 == 0)
                    {
                        await Task.Yield();
                    }
                }

                // =====================================================
                // LƯU JSON
                // =====================================================

                progressForm.UpdateProgress(
                    total,
                    total,
                    "Đang cập nhật JSON...");

                bool daLuu = await Task.Run(() => LuuCapNhatRecordsVaoJson(recordsToProcess));

                progressForm.UpdateProgress(total, total, "Đang cập nhật giao diện...", "Chuyển đổi file PDF");

                if (selectedDangVienRecords != null &&
                    selectedDangVienRecords.Count > 0)
                {
                    HienThiDanhSachPdf(selectedDangVienRecords);
                }
                else
                {
                    HienThiTree(_allRecords);
                }
                // =====================================================
                // HOÀN THÀNH
                // =====================================================

                progressForm.UpdateProgress(
                    total,
                    total,
                    "Đã hoàn thành", "Chuyển đổi file PDF");

                await Task.Delay(300);

                // =====================================================
                // THÔNG BÁO
                // =====================================================

                MessageBox.Show(
                    $"Đã lưu file hoàn thành.\n\n" +
                    $"Tổng bản ghi: {total:N0}\n" +
                    $"Đã chuyển: {daChuyen:N0}\n" +
                    $"Không tìm thấy: {khongTonTai:N0}\n" +
                    $"File đã tồn tại: {biTrung:N0}\n" +
                    $"Lỗi / thiếu dữ liệu: {loi:N0}\n" +
                    $"Cập nhật JSON: {(daLuu ? "Thành công" : "Thất bại")}",
                    "Hoàn thành",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Có lỗi khi lưu file:\n\n" +
                    ex.Message,
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                // =================================================
                // ĐÓNG PROGRESS FORM
                // =================================================

                if (progressForm != null &&
                    !progressForm.IsDisposed)
                {
                    progressForm.Close();
                    progressForm.Dispose();
                }

                btnSave.Enabled = true;
            }
        }
        private List<FrmChonTenFile.DanhMucTaiLieu> LayDanhSachTenFile()
        {
            try
            {
                if (!File.Exists(danhMucJsonPath))
                {
                    MessageBox.Show(
                        "Không tìm thấy file danh mục tài liệu:\n\n" +
                        danhMucJsonPath,
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return new List<FrmChonTenFile.DanhMucTaiLieu>();
                }

                string json = File.ReadAllText(
                    danhMucJsonPath,
                    Encoding.UTF8);

                List<FrmChonTenFile.DanhMucTaiLieu>? danhSach = System.Text.Json.JsonSerializer.Deserialize<
                        List<FrmChonTenFile.DanhMucTaiLieu>>(
                            json,
                            new JsonSerializerOptions
                            {
                                PropertyNameCaseInsensitive = true
                            });

                if (danhSach == null)
                    return new List<FrmChonTenFile.DanhMucTaiLieu>();

                return danhSach
                    .Where(x =>
                        !string.IsNullOrWhiteSpace(x.ten_file))
                    .OrderBy(x =>
                    {
                        if (int.TryParse(x.tt, out int so))
                            return so;

                        return int.MaxValue;
                    })
                    .ToList();
            }
            catch (System.Text.Json.JsonException ex)
            {
                MessageBox.Show(
                    "File danh mục tài liệu không đúng định dạng JSON:\n\n" +
                    ex.Message,
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                return new List<FrmChonTenFile.DanhMucTaiLieu>();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Không thể đọc danh mục tài liệu:\n\n" +
                    ex.Message,
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                return new List<FrmChonTenFile.DanhMucTaiLieu>();
            }
        }
    }

}

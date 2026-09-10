using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace renamePDF_v2
{
    public partial class FrmChonTenFile : Form
    {
        // ==========================================================
        // MODEL DANH MỤC
        // ==========================================================

        public class DanhMucTaiLieu
        {
            public string tt { get; set; } = "";
            public string ten_tai_lieu { get; set; } = "";
            public string ten_file { get; set; } = "";
            public string do_uu_tien { get; set; } = "";
        }

        // ==========================================================
        // DỮ LIỆU
        // ==========================================================

        private readonly List<DanhMucTaiLieu> _danhSachGoc;

        private List<DanhMucTaiLieu> _danhSachHienThi = new();

        // Tên hiện tại của record
        private readonly string _tenHienTai;

        // Ngăn sự kiện lồng nhau
        private bool _dangXuLy = false;

        // ==========================================================
        // KẾT QUẢ TRẢ VỀ
        // ==========================================================

        /// <summary>
        /// Chỉ trả về ten_file.
        /// </summary>
        public string TenDuocChon { get; private set; } = "";

        // ==========================================================
        // CONSTRUCTOR
        // ==========================================================

        public FrmChonTenFile(string tenHienTai,List<DanhMucTaiLieu> danhSachTenFile)
        {
            InitializeComponent();

            _tenHienTai =tenHienTai?.Trim() ?? "";

            _danhSachGoc = danhSachTenFile.Where(x =>!string.IsNullOrWhiteSpace(x.ten_file)).OrderBy(x => LaySoTT(x.tt)).ToList();
            // ======================================================
            // SỰ KIỆN
            // ======================================================
            txtTimKiem.TextChanged +=TxtTimKiem_TextChanged;
            cboTenFile.SelectedIndexChanged +=CboTenFile_SelectedIndexChanged;
            btnChon.Click +=BtnChon_Click;
            // ======================================================
            // HIỂN THỊ DANH SÁCH BAN ĐẦU
            // ======================================================
            HienThiDanhSach(_danhSachGoc);
            // ======================================================
            // HIỂN THỊ TÊN HIỆN TẠI
            //
            // KHÔNG đưa tên hiện tại vào ô tìm kiếm
            // ======================================================
            HienThiTenHienTai();
            // ======================================================
            // Ô TÌM KIẾM LUÔN RỖNG
            // ======================================================
            _dangXuLy = true;
            txtTimKiem.Clear();
            _dangXuLy = false;
            txtTimKiem.Focus();
        }

        // ==========================================================
        // LẤY SỐ TT
        // ==========================================================

        private static int LaySoTT(string tt)
        {
            if (int.TryParse(tt, out int so))return so;
            return int.MaxValue;
        }

        // ==========================================================
        // HIỂN THỊ DANH SÁCH COMBOBOX
        // ==========================================================

        private void HienThiDanhSach(IEnumerable<DanhMucTaiLieu> danhSach)
        {
            _danhSachHienThi =danhSach.ToList();
            cboTenFile.BeginUpdate();
            try
            {
                cboTenFile.Items.Clear();
                foreach (DanhMucTaiLieu item in _danhSachHienThi)
                {
                    string hienThi =$"{item.tt} | {item.ten_file}";
                    cboTenFile.Items.Add(hienThi);
                }
            }
            finally
            {
                cboTenFile.EndUpdate();
            }
        }

        // ==========================================================
        // HIỂN THỊ TÊN HIỆN TẠI
        // ==========================================================

        private void HienThiTenHienTai()
        {
            if (string.IsNullOrWhiteSpace(_tenHienTai))
            {
                cboTenFile.SelectedIndex = -1;
                cboTenFile.Text = "";

                return;
            }

            // ------------------------------------------------------
            // Tìm trong danh mục
            // ------------------------------------------------------

            DanhMucTaiLieu? item =
                _danhSachGoc.FirstOrDefault(x =>
                    string.Equals(
                        x.ten_file.Trim(),
                        _tenHienTai,
                        StringComparison.OrdinalIgnoreCase));

            // ------------------------------------------------------
            // CÓ TRONG DANH MỤC
            // ------------------------------------------------------

            if (item != null)
            {
                int index =
                    _danhSachHienThi.IndexOf(item);

                if (index >= 0)
                {
                    _dangXuLy = true;

                    try
                    {
                        cboTenFile.SelectedIndex =
                            index;
                    }
                    finally
                    {
                        _dangXuLy = false;
                    }

                    return;
                }
            }

            // ------------------------------------------------------
            // KHÔNG CÓ TRONG DANH MỤC
            //
            // Vẫn hiển thị tên cũ
            // ------------------------------------------------------

            _dangXuLy = true;

            try
            {
                cboTenFile.SelectedIndex = -1;

                cboTenFile.Text =
                    _tenHienTai;
            }
            finally
            {
                _dangXuLy = false;
            }
        }

        // ==========================================================
        // TÌM KIẾM
        // ==========================================================

        private void TxtTimKiem_TextChanged(
            object? sender,
            EventArgs e)
        {
            if (_dangXuLy)
                return;

            string tuKhoa =
                txtTimKiem.Text.Trim();

            // ------------------------------------------------------
            // Không có từ khóa
            // ------------------------------------------------------

            if (string.IsNullOrWhiteSpace(tuKhoa))
            {
                HienThiDanhSach(
                    _danhSachGoc);

                return;
            }

            // ------------------------------------------------------
            // Tìm theo:
            //
            // 1. tt
            // 2. ten_tai_lieu
            // 3. ten_file
            // ------------------------------------------------------

            List<DanhMucTaiLieu> ketQua =
                _danhSachGoc
                    .Where(x =>
                        ChuaTuKhoa(
                            x.tt,
                            tuKhoa)
                        ||
                        ChuaTuKhoa(
                            x.ten_tai_lieu,
                            tuKhoa)
                        ||
                        ChuaTuKhoa(
                            x.ten_file,
                            tuKhoa))
                    .ToList();

            HienThiDanhSach(ketQua);

            // ------------------------------------------------------
            // Xóa lựa chọn hiện tại
            // ------------------------------------------------------

            _dangXuLy = true;

            try
            {
                cboTenFile.SelectedIndex = -1;
                cboTenFile.Text = tuKhoa;

                cboTenFile.SelectionStart =
                    cboTenFile.Text.Length;
            }
            finally
            {
                _dangXuLy = false;
            }

            // ------------------------------------------------------
            // Mở danh sách kết quả
            // ------------------------------------------------------

            if (ketQua.Count > 0)
            {
                cboTenFile.DroppedDown = true;
            }
        }

        // ==========================================================
        // TÌM KIẾM KHÔNG PHÂN BIỆT HOA THƯỜNG
        // ==========================================================

        private static bool ChuaTuKhoa(
            string? text,
            string tuKhoa)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            return text.Contains(
                tuKhoa,
                StringComparison.OrdinalIgnoreCase);
        }

        // ==========================================================
        // CHỌN ITEM TRONG COMBOBOX
        // ==========================================================

        private void CboTenFile_SelectedIndexChanged(
            object? sender,
            EventArgs e)
        {
            if (_dangXuLy)
                return;

            int index =
                cboTenFile.SelectedIndex;

            if (index < 0 ||
                index >= _danhSachHienThi.Count)
            {
                return;
            }

            DanhMucTaiLieu item =
                _danhSachHienThi[index];

            string hienThi =
                $"{item.tt} | {item.ten_file}";

            _dangXuLy = true;

            try
            {
                // --------------------------------------------------
                // ComboBox hiển thị:
                //
                // 01 | 01.Ly lich cua nguoi xin vao dang.pdf
                // --------------------------------------------------

                cboTenFile.Text =
                    hienThi;

                // --------------------------------------------------
                // Ô tìm kiếm trở về tên file
                // --------------------------------------------------

                txtTimKiem.Text =
                    item.ten_file;

                txtTimKiem.SelectionStart =
                    txtTimKiem.Text.Length;
            }
            finally
            {
                _dangXuLy = false;
            }
        }

        // ==========================================================
        // NÚT CHỌN
        // ==========================================================

        private void BtnChon_Click(
            object? sender,
            EventArgs e)
        {
            int index =
                cboTenFile.SelectedIndex;

            // ======================================================
            // NGƯỜI DÙNG CHỌN TRONG DANH SÁCH
            // ======================================================

            if (index >= 0 &&
                index < _danhSachHienThi.Count)
            {
                DanhMucTaiLieu item =
                    _danhSachHienThi[index];

                // Chỉ trả về ten_file
                TenDuocChon =
                    item.ten_file.Trim();

                DialogResult =
                    DialogResult.OK;

                Close();

                return;
            }

            // ======================================================
            // TRƯỜNG HỢP TÊN CŨ KHÔNG CÓ TRONG DANH MỤC
            // ======================================================

            if (!string.IsNullOrWhiteSpace(
                _tenHienTai))
            {
                if (string.Equals(
                    cboTenFile.Text.Trim(),
                    _tenHienTai,
                    StringComparison.OrdinalIgnoreCase))
                {
                    TenDuocChon =
                        _tenHienTai;

                    DialogResult =
                        DialogResult.OK;

                    Close();

                    return;
                }
            }

            // ======================================================
            // CHƯA CHỌN
            // ======================================================

            MessageBox.Show(
                "Vui lòng chọn một tên tài liệu trong danh sách.",
                "Thông báo",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }
}
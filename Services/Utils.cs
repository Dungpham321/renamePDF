using renamePDF_v2.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace renamePDF_v2.Services
{
    public static class Utils
    {
        private static readonly string settingsJsonPath = Path.Combine(Application.StartupPath, "data", "settings.json");
        public static string NormalizeKey(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            // Bỏ dấu tiếng Việt
            string normalized = text.Normalize(
                System.Text.NormalizationForm.FormD);

            StringBuilder sb = new StringBuilder();

            foreach (char c in normalized)
            {
                var category = CharUnicodeInfo.GetUnicodeCategory(c);

                if (category != UnicodeCategory.NonSpacingMark)
                {
                    sb.Append(c);
                }
            }

            string result = sb.ToString()
                .Normalize(System.Text.NormalizationForm.FormC);

            // Đổi đ/ch
            result = result
                .Replace("đ", "d")
                .Replace("Đ", "D");

            // Các ký tự không phải chữ/số chuyển thành _
            result = Regex.Replace(
                result,
                @"[^a-zA-Z0-9]+",
                "_");

            // Chuyển PascalCase / camelCase thành snake_case
            result = Regex.Replace(
                result,
                @"([a-z0-9])([A-Z])",
                "$1_$2");

            // Xóa _ đầu/cuối
            result = result.Trim('_');

            return result.ToLowerInvariant();
        }

        public static ParsedFolder ParseFolder(string path)
        {
            ParsedFolder result = new ParsedFolder();

            if (string.IsNullOrWhiteSpace(path))
                return result;

            path = path.Replace('/', '\\');

            // ============================================================
            // LẤY TÊN FILE PDF
            // ============================================================

            result.TenFile = Path.GetFileName(path).Trim();

            string[] folders =
                path.Split(
                    new[] { '\\' },
                    StringSplitOptions.RemoveEmptyEntries);

            // ============================================================
            // TÌM CHI BỘ
            // ============================================================

            int chiBoIndex = -1;

            for (int i = 0;
                 i < folders.Length;
                 i++)
            {
                if (Regex.IsMatch(
                    folders[i],
                    @"^\d+_ChiBo",
                    RegexOptions.IgnoreCase))
                {
                    chiBoIndex = i;
                    break;
                }
            }

            if (chiBoIndex == -1)
                return result;

            // ============================================================
            // TÊN CHI BỘ
            // ============================================================

            string chiBoFolder =
                folders[chiBoIndex];

            int underscore =
                chiBoFolder.IndexOf('_');

            if (underscore >= 0)
            {
                result.TenChiBo =
                    chiBoFolder
                        .Substring(underscore + 1)
                        .Trim();
            }

            // ============================================================
            // TÊN ĐẢNG VIÊN + NĂM SINH
            // ============================================================

            int dangVienIndex =
                chiBoIndex + 1;

            if (dangVienIndex >= folders.Length)
                return result;

            string dangVienFolder =
                folders[dangVienIndex];

            Match match =
                Regex.Match(
                    dangVienFolder,
                    @"^\d+_(.+)_(\d{4})$");

            if (match.Success)
            {
                string tenDangVien =
                    match.Groups[1]
                        .Value
                        .Trim();

                string namSinh =
                    match.Groups[2]
                        .Value
                        .Trim();

                result.DangVienNamSinh =
                    $"{tenDangVien} - {namSinh}";
            }

            // ============================================================
            // NHÓM VĂN BẢN
            // ============================================================

            int groupIndex =
                dangVienIndex + 1;

            if (groupIndex < folders.Length)
            {
                string groupFolder =
                    folders[groupIndex];

                if (Regex.IsMatch(
                    groupFolder,
                    @"^\d+_"))
                {
                    result.NhomVanBan =
                        groupFolder.Trim();
                }
            }

            return result;
        }

        public static string LamSachTenThuMuc(string ten)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                ten = ten.Replace(c, '_');
            }

            return ten.Trim();
        }

        public static string LayTenDangVien(string dangVienNamSinh)
        {
            if (string.IsNullOrWhiteSpace(dangVienNamSinh))
                return "";

            int index = dangVienNamSinh.IndexOf(" - ");

            if (index >= 0)
            {
                return dangVienNamSinh
                    .Substring(0, index)
                    .Trim();
            }

            return dangVienNamSinh.Trim();
        }
        public static string LayThuMucTuSetting()
        {
            string settingPath = settingsJsonPath;
            if (!File.Exists(settingPath))
            {
                throw new Exception(
                    $"Không tìm thấy file setting.json:\n{settingPath}");
            }

            string json =
                File.ReadAllText(
                    settingPath,
                    Encoding.UTF8);

            using JsonDocument doc =
                JsonDocument.Parse(json);

            if (!doc.RootElement.TryGetProperty(
                    "ThuMuc",
                    out JsonElement element))
            {
                throw new Exception(
                    "setting.json không có thuộc tính \"ThuMuc\".");
            }

            string path = element.GetString()?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(path))
            {
                throw new Exception(
                    "ThuMuc trong setting.json đang trống.");
            }

            return path;
        }

        public static string TaoTenFileTrung(string tenFile, int stt)
        {
            string extension = Path.GetExtension(tenFile);

            string tenKhongDuoi = Path.GetFileNameWithoutExtension(tenFile);

            return $"{tenKhongDuoi} ({stt}){extension}";
        }

    }
}

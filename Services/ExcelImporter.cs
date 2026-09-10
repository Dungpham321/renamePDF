using ClosedXML.Excel;
using renamePDF_v2.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace renamePDF_v2.Services
{
    public class ExcelImporter
    {

        public List<Dictionary<string, object?>> ImportExcel(string filePath, Action<int, int, string>? progress = null)
        {
            List<Dictionary<string, object?>> result =
                new List<Dictionary<string, object?>>();

            using XLWorkbook workbook = new XLWorkbook(filePath);
            IXLWorksheet worksheet = workbook.Worksheets.First();
            IXLRow? headerRow = worksheet.FirstRowUsed();

            if (headerRow == null)
            {
                throw new Exception(
                    "Không tìm thấy dòng tiêu đề.");
            }

            int firstColumn =
                headerRow.FirstCellUsed()
                .Address.ColumnNumber;

            int lastColumn =
                headerRow.LastCellUsed()
                .Address.ColumnNumber;

            List<string> columns =
                GetExcelColumns(
                    worksheet,
                    headerRow,
                    firstColumn,
                    lastColumn);

            int pathIndex = FindColumn(columns, "duong_dan");

            if (pathIndex == -1)
            {
                pathIndex = FindColumn(columns, "FOLDER");
            }

            if (pathIndex == -1)
            {
                throw new Exception(
                    "Không tìm thấy cột 'Đường dẫn' hoặc 'FOLDER'.");
            }

            int startRow =
                headerRow.RowNumber() + 1;

            int lastRow =
                worksheet.LastRowUsed()?.RowNumber()
                ?? startRow - 1;

            // Tổng số dòng cần đọc
            int totalRows =
                Math.Max(0, lastRow - startRow + 1);

            int currentRow = 0;

            for (int rowNumber = startRow; rowNumber <= lastRow; rowNumber++)
            {
                currentRow++;

                // Lấy tên file đang xử lý
                string currentFileName = "";

                Dictionary<string, object?> item = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);

                bool hasData = false;

                for (int i = 0; i < columns.Count; i++)
                {
                    int excelColumn =
                        firstColumn + i;

                    IXLCell cell =
                        worksheet.Cell(
                            rowNumber,
                            excelColumn);

                    object? value = GetCellValue(cell);

                    item[columns[i]] = value;

                    // Nếu đang đọc cột đường dẫn thì lấy tên file
                    if (i == pathIndex && value != null)
                    {
                        currentFileName =
                            Path.GetFileName(
                                value.ToString() ?? "");
                    }

                    if (value != null &&
                        !string.IsNullOrWhiteSpace(
                            value.ToString()))
                    {
                        hasData = true;
                    }
                }

                // Báo tiến độ
                progress?.Invoke(
                    currentRow,
                    totalRows,
                    currentFileName);

                if (!hasData)
                    continue;

                string path =
                    item[columns[pathIndex]]
                    ?.ToString()
                    ?.Trim()
                    ?? "";

                ParsedFolder parsed = Utils.ParseFolder(path);
                item["TenChiBo"] = parsed.TenChiBo;
                item["DangVienNamSinh"] = parsed.DangVienNamSinh;
                item["NhomVanBan"] = parsed.NhomVanBan;
                // Ưu tiên lấy từ cột ten_file
                string tenFile = "";

                // Tìm cột ten_file trong Excel
                int tenFileIndex = FindColumn(columns, "ten_file");

                if (tenFileIndex >= 0)
                {
                    object? valueTenFile = item[columns[tenFileIndex]];

                    tenFile = valueTenFile?.ToString()?.Trim() ?? "";
                }

                // Nếu ten_file không có dữ liệu thì mới lấy từ đường dẫn
                if (string.IsNullOrWhiteSpace(tenFile))
                {
                    tenFile = parsed.TenFile;
                }

                item["TenFile"] = tenFile;

                result.Add(item);
            }

            return result;
        }
        public List<string> GetExcelColumns(IXLWorksheet worksheet, IXLRow headerRow, int firstColumn, int lastColumn)
        {
            List<string> columns =
                new List<string>();

            for (int col = firstColumn;
                 col <= lastColumn;
                 col++)
            {
                string columnName =
                    worksheet
                    .Cell(headerRow.RowNumber(), col)
                    .GetString()
                    .Trim();

                string key = Utils.NormalizeKey(columnName);

                if (string.IsNullOrWhiteSpace(key))
                {
                    key = $"column_{col}";
                }

                string originalKey = key;
                int suffix = 2;

                while (columns.Contains(key))
                {
                    key = $"{originalKey}_{suffix}";
                    suffix++;
                }

                columns.Add(key);
            }

            return columns;
        }

        public int FindColumn(List<string> columns, string columnName)
        {
            for (int i = 0; i < columns.Count; i++)
            {
                if (string.Equals(
                    columns[i].Trim(),
                    columnName.Trim(),
                    StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }
        public object? GetCellValue(IXLCell cell)
        {
            if (cell.IsEmpty())
                return null;

            string value =
                cell.GetString();

            if (string.IsNullOrWhiteSpace(value))
                return null;

            return value;
        }

    }
}

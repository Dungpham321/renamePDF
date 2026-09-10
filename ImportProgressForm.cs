using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace renamePDF_v2
{
    public partial class ImportProgressForm : Form
    {
        public ImportProgressForm()
        {
            InitializeComponent();

            StartPosition =
                FormStartPosition.CenterParent;

            FormBorderStyle =
                FormBorderStyle.FixedDialog;

            ControlBox = false;
            MaximizeBox = false;
            MinimizeBox = false;

            lblTitle.Text =
                "ĐANG IMPORT EXCEL";

            lblFile.Text =
                "Đang chuẩn bị...";

            lblCount.Text =
                "0 / 0";

            progressBar.Minimum = 0;
            progressBar.Maximum = 100;
            progressBar.Value = 0;
        }

        public void UpdateProgress(int current,int total,string fileName, string Title = "ĐANG XỬ LÝ")
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() =>
                    UpdateProgress(
                        current,
                        total,
                        fileName, Title)));

                return;
            }

            if (total <= 0) total = 1;

            int percent = (int)((double)current / total * 100);

            percent = Math.Max(0, Math.Min(100, percent));
            progressBar.Value = percent;

            lblTitle.Text = Title;
            lblCount.Text =
                $"{current:N0} / {total:N0}";

            lblFile.Text =
                string.IsNullOrWhiteSpace(fileName)
                    ? "Đang đọc dữ liệu..."
                    : $"Đang đọc: {fileName}";

            Refresh();
        }

    }
}

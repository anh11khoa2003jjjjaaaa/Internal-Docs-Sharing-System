using Guna.UI2.WinForms;
using OfficeOpenXml;
using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ChiaseNoiBo
{
    public partial class Form3_LoadExcel : Form
    {
        private readonly string fileId;
        private readonly string fileName;
        private string tempFilePath;
        private readonly Form1 _form1;
        private Panel loadingOverlay;
        private Label loadingLabel;
        private ProgressBar progressBar; // hoặc PictureBox dùng ảnh gif loading

        public Form3_LoadExcel(string fileId, string fileName)
        {
            InitializeComponent();
            this.fileId = fileId;
            this.fileName = fileName;
            this.Text = $"Xem File: {fileName}";
            _form1 = new Form1(fileId, fileName);
        }

        private async void Form3_LoadExcel_Load(object sender, EventArgs e)
        {
            
            tempFilePath = Path.Combine(Path.GetTempPath(), fileName);
            guna2TabControl1.Visible = false;
            //ShowLoading();
            try
            {
                using (var stream = await _form1.DownloadFileFromGoogleDriveAsync(fileId))
                {
                    if (stream == null || stream.Length < 100)
                        throw new Exception("File không hợp lệ hoặc quá nhỏ!");

                    using (var fileStream = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write))
                    {
                        stream.CopyTo(fileStream);
                    }
                }
                guna2TabControl1.TabMenuVisible = false;

                await ShowPreviewAsync(fileName, tempFilePath);
                //HideLoading();
                guna2TabControl1.Visible = true;

            }
            catch (Exception ex)
            {
                _form1.ShowError($"Lỗi khi xử lý file: {ex.Message}");
            }
        }

        private async Task ShowPreviewAsync(string fileName, string filePath)
        {
            string ext = Path.GetExtension(fileName).ToLower();

            guna2TabControl1.Visible = false;
            // Xoá toàn bộ tab cũ (nếu có)

            if (ext == ".xlsx" || ext == ".xls" || ext == ".csv")
            {
                guna2TabControl1.Visible = true;
                await LoadExcelToTabControlAsync(filePath);
            }
            else
            {
                _form1.ShowError("Định dạng file không được hỗ trợ.");
            }
        }

        private async Task LoadExcelToTabControlAsync(string filePath)
        {
            guna2TabControl1.TabMenuVisible = true;
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    if (package.Workbook.Worksheets.Count == 0)
                        throw new Exception("Không tìm thấy dữ liệu trong Excel!");

                    guna2TabControl1.TabPages.Clear(); // Xóa tab cũ nếu có

                    foreach (var worksheet in package.Workbook.Worksheets)
                    {
                        if (worksheet.Dimension == null)
                            continue;

                        var dt = ConvertWorksheetToDataTable(worksheet);

                        if (dt.Rows.Count > 0 || dt.Columns.Count > 0)
                        {
                            AddTabWithDataTable(worksheet.Name, dt);
                        }
                    }

                    if (guna2TabControl1.TabPages.Count == 0)
                        throw new Exception("Tất cả các sheet đều trống!");

                    guna2TabControl1.SelectedIndex = 0; 
                }
            }
            catch (Exception ex)
            {
                _form1.ShowError($"Lỗi khi đọc dữ liệu Excel: {ex.Message}");
            }
        }

        //private void AddTabWithDataTable(string tabName, DataTable dt)
        //{
        //    var tabPage = new TabPage(tabName);

        //    var dgv = new DataGridView
        //    {
        //        Dock = DockStyle.Fill,
        //        ReadOnly = true,
        //        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
        //        DataSource = dt,
        //        BorderStyle = BorderStyle.None,
        //        Font = new Font("Segoe UI", 10), // Cỡ chữ mặc định cho toàn bộ lưới
        //        ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
        //        {
        //            Font = new Font("Segoe UI", 10, FontStyle.Bold), // Header đậm
        //            BackColor = Color.LightGray,
        //            ForeColor = Color.Black,
        //            Alignment = DataGridViewContentAlignment.MiddleCenter
        //        },
        //        EnableHeadersVisualStyles = false // Cho phép style tùy chỉnh header hoạt động
        //    };


        //    tabPage.Controls.Add(dgv);
        //    guna2TabControl1.TabPages.Add(tabPage);
        //}
        private void AddTabWithDataTable(string tabName, DataTable dt)
        {
            var tabPage = new TabPage(tabName);

            var dgv = new Guna.UI2.WinForms.Guna2DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                DataSource = dt,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 10),
                EnableHeadersVisualStyles = false,
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    Font = new Font("Segoe UI", 10, FontStyle.Bold),
                    BackColor = Color.DeepSkyBlue,
                    ForeColor = Color.Black,
                   
                    Alignment = DataGridViewContentAlignment.MiddleCenter
                },
                ColumnHeadersHeight = 40,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,

                RowTemplate = { Height = 35 }, // Tuỳ chỉnh chiều cao dòng nếu muốn
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(245, 245, 245)
                },
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
            };

            tabPage.Controls.Add(dgv);
            guna2TabControl1.TabPages.Add(tabPage);
        }


        private DataTable ConvertWorksheetToDataTable(ExcelWorksheet worksheet)
        {
            var dt = new DataTable();

            if (worksheet.Dimension == null)
                return dt;

            int startCol = worksheet.Dimension.Start.Column;
            int endCol = worksheet.Dimension.End.Column;
            int startRow = worksheet.Dimension.Start.Row;
            int endRow = worksheet.Dimension.End.Row;
            
            for (int col = startCol; col <= endCol; col++)
            {
                string colName = worksheet.Cells[startRow, col].Text.Trim();
                if (string.IsNullOrWhiteSpace(colName))
                    colName = $"Column{col}";
                dt.Columns.Add(colName);
            }
            //Bỏ qua dòng đầu tiên
            for (int row = startRow + 1; row <= endRow; row++)
            {
                var newRow = dt.NewRow();
                for (int col = startCol; col <= endCol; col++)
                {
                    newRow[col - startCol] = worksheet.Cells[row, col].Text;
                }
                dt.Rows.Add(newRow);
            }

            return dt;
        }
        

    }
}

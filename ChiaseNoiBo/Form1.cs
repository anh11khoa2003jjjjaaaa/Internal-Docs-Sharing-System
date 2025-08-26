using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Apis.Drive.v3;
using Guna.UI2.WinForms;
using OfficeOpenXml;
using PdfiumViewer;
using ChiaseNoiBo.Helpers;

namespace ChiaseNoiBo
{
    public partial class Form1 : Form
    {
        private readonly string fileId;
        private readonly string fileName;
        private string tempFilePath;
        private PdfViewer pdfViewer;
        private int zoomLevel = 100;
        public Form1(string fileId, string fileName)
        {
            InitializeComponent();
            this.fileId = fileId;
            this.fileName = fileName;
            this.Text = $"Xem File: {fileName}";
        }

        private async void Form1_Load(object sender, EventArgs e)
        {
            tempFilePath = Path.Combine(Path.GetTempPath(), fileName);

            try
            {
                using (var stream = await DownloadFileFromGoogleDriveAsync(fileId))
                {
                    if (stream == null || stream.Length < 100)
                        throw new Exception("File không hợp lệ hoặc quá nhỏ!");

                    using (var fileStream = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write))
                    {
                        stream.CopyTo(fileStream);
                    }
                }

                ShowPreview(fileName, tempFilePath);
            }
            catch (Exception ex)
            {
                ShowError($"Lỗi khi xử lý file: {ex.Message}");
            }
        }

        public async Task<MemoryStream> DownloadFileFromGoogleDriveAsync(string fileId)
        {
            try
            {
                var service = GoogleDriveHelper.GetDriveService();
                var request = service.Files.Get(fileId);

                var stream = new MemoryStream();
                await request.DownloadAsync(stream);
                stream.Position = 0;
                return stream;
            }
            catch (Exception ex)
            {
                ShowError($"Lỗi khi tải file từ Google Drive: {ex.Message}");
                return null;
            }
        }

        public async Task LoadExcelToGridAsync(string filePath)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    if (package.Workbook.Worksheets.Count == 0)
                        throw new Exception("Không tìm thấy dữ liệu trong Excel!");

                    var worksheet = package.Workbook.Worksheets[0];
                    var dt = ConvertWorksheetToDataTable(worksheet);
                    guna2DataGridView1.DataSource = dt;
                }
            }
            catch (Exception ex)
            {
                ShowError($"Lỗi khi đọc dữ liệu Excel: {ex.Message}");
            }
        }

        private DataTable ConvertWorksheetToDataTable(ExcelWorksheet worksheet)
        {
            var dt = new DataTable();

            for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
            {
                dt.Columns.Add(worksheet.Cells[1, col].Text);
            }

            for (int row = worksheet.Dimension.Start.Row + 1; row <= worksheet.Dimension.End.Row; row++)
            {
                var newRow = dt.NewRow();
                for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                {
                    newRow[col - 1] = worksheet.Cells[row, col].Text;
                }
                dt.Rows.Add(newRow);
            }

            return dt;
        }

        public void ShowError(string message)
        {
            guna2MessageDialog1.Icon = MessageDialogIcon.Error;
            guna2MessageDialog1.Show(message, "Lỗi");
        }

        public void ShowPreview(string fileName, string filePath)
        {
            string ext = Path.GetExtension(fileName).ToLower();

            printPreviewControl1.Controls.Clear();
            guna2DataGridView1.Visible = false;
            printPreviewControl1.Visible = true;

            if (ext == ".pdf")
            {
                ShowPdfWithPdfium(filePath);
            }
            else if (ext == ".xlsx" || ext == ".xls")
            {
                printPreviewControl1.Visible = false;
                guna2DataGridView1.Visible = true;
                _ = LoadExcelToGridAsync(filePath);
            }
            else if (ext == ".txt")
            {
                var txtBox = new TextBox
                {
                    Dock = DockStyle.Fill,
                    Multiline = true,
                    ScrollBars = ScrollBars.Both,
                    Text = File.ReadAllText(filePath)
                };
                printPreviewControl1.Controls.Add(txtBox);
            }
            else if (ext == ".doc" || ext == ".docx")
            {
                ShowWordDocument(filePath);
            }
            else
            {
                ShowError("Không hỗ trợ xem trước định dạng file này.");
            }
        }
        private void ShowWordDocument(string filePath)
        {
            try
            {
                var wordApp = new Microsoft.Office.Interop.Word.Application();
                var wordDoc = wordApp.Documents.Open(filePath);

                // Đặt chế độ hiển thị của Word dưới dạng chế độ xem Print Preview
                wordApp.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView;

                // Hiển thị Word trong Form
                wordApp.Visible = true;
            }
            catch (Exception ex)
            {
                ShowError($"Lỗi khi mở file Word: {ex.Message}");
            }
        } 

            private void ShowPdfWithPdfium(string filePath)
        {
            try
            {
                // Khởi tạo lại PdfViewer
                pdfViewer = new PdfViewer()
                {
                    Dock = DockStyle.Fill,
                    ShowToolbar = false,
                };

                // Load tài liệu PDF
                var document = PdfiumViewer.PdfDocument.Load(filePath);
                pdfViewer.Document = document;

                // Tạo panel chứa các nút điều khiển
                var panelControl = new FlowLayoutPanel()
                {
                    Height = 40,
                    Dock = DockStyle.Top,
                    FlowDirection = FlowDirection.LeftToRight,
                    Padding = new Padding(5),
                    BackColor = Color.LightGray
                };

                // Tạo các nút điều khiển
                var btnZoomIn = new Button() { Text = "Phóng to", Width = 80 };
                var btnZoomOut = new Button() { Text = "Thu nhỏ", Width = 80 };
             

                // Sự kiện ZoomIn: Chuyển sang chế độ FitWidth (mô phỏng zoom)
                // Sự kiện ZoomIn: Phóng to theo chiều rộng (FitWidth)
                btnZoomIn.Click += (s, e) =>
                {
                    if (pdfViewer.Document != null)
                    {
                        // Đặt chế độ phóng to vừa vặn theo chiều rộng
                        pdfViewer.ZoomMode = PdfViewerZoomMode.FitWidth;
                    }
                };

                // Sự kiện ZoomOut: Thu nhỏ theo chiều cao (FitHeight)
                btnZoomOut.Click += (s, e) =>
                {
                    if (pdfViewer.Document != null)
                    {
                        // Đặt chế độ thu nhỏ vừa vặn theo chiều cao
                        pdfViewer.ZoomMode = PdfViewerZoomMode.FitHeight;
                    }
                };


                

                // Thêm các nút vào panel
                panelControl.Controls.AddRange(new Control[] { btnZoomIn, btnZoomOut });

                // Xóa và thêm mới control
                printPreviewControl1.Controls.Clear();
                printPreviewControl1.Controls.Add(panelControl);
                printPreviewControl1.Controls.Add(pdfViewer);
            }
            catch (Exception ex)
            {
                ShowError($"Lỗi khi mở file PDF: {ex.Message}");
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            try
            {
                if (File.Exists(tempFilePath))
                    File.Delete(tempFilePath);
            }
            catch { }
        }
    }
}


using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Wordprocessing;
namespace ChiaseNoiBo
{
    public partial class DeXuatTangLuongControl : UserControl
    {
        public DeXuatTangLuongControl()
        {
            InitializeComponent();
        }
        public string Lanhdao => txt_lanhdao.Text;
        public string HoTen => txt_name.Text;
        public string Phongban => txt_bophan.Text;
        public string Vitri => txt_chucvu.Text;
        public string LyDo => richTextBox1_lydo.Text;
        public double Salary_current => (double)guna2NumericUpDown1_hientai.Value;
        public double Salary_expected => (double)guna2NumericUpDown2_mongmuon.Value;
        public string TuNgay => guna2DateTimePicker1.Value.ToShortDateString();

       
        public static async Task<string> DownloadFileFromGoogleDriveAsync(string fileId, string savePath)
        {
            var url = $"https://drive.google.com/uc?export=download&id={fileId}";
            using (var client = new HttpClient())
            {
                var response = await client.GetAsync(url);
                if (response.IsSuccessStatusCode)
                {
                    using (var fs = new FileStream(savePath, FileMode.Create, FileAccess.Write))
                    {
                        await response.Content.CopyToAsync(fs);
                    }
                    return savePath;
                }
                else
                {
                    throw new Exception("Tải file từ Google Drive thất bại.");
                }
            }
        }
        public static void ReplacePlaceholders(string filePath, Dictionary<string, string> replacements)
        {
            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
                {
                    if (wordDoc.MainDocumentPart == null)
                    {
                        throw new Exception("File Word không hợp lệ: thiếu MainDocumentPart");
                    }

                    var body = wordDoc.MainDocumentPart.Document.Body;
                    if (body == null)
                    {
                        throw new Exception("File Word không hợp lệ: thiếu Body");
                    }

                    bool foundAnyPlaceholder = false;

                    foreach (var text in body.Descendants<Text>())
                    {
                        foreach (var kvp in replacements)
                        {
                            if (text.Text.Contains(kvp.Key))
                            {
                                text.Text = text.Text.Replace(kvp.Key, kvp.Value);
                                foundAnyPlaceholder = true;
                                Console.WriteLine($"Đã thay thế: {kvp.Key} => {kvp.Value}");
                            }
                        }
                    }

                    if (!foundAnyPlaceholder)
                    {
                        throw new Exception("Không tìm thấy placeholder nào trong file");
                    }

                    wordDoc.MainDocumentPart.Document.Save();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi khi thay thế placeholder: {ex.Message}", ex);
            }
        }
        //public static void ReplacePlaceholders(string filePath, Dictionary<string, string> replacements)
        //{
        //    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
        //    {
        //        var body = wordDoc.MainDocumentPart.Document.Body;
        //        foreach (var text in body.Descendants<Text>())
        //        {
        //            Console.WriteLine("Dữ liệu thay thế:");
        //            foreach (var kvp in replacements)
        //            {
        //                if (text.Text.Contains(kvp.Key))
        //                {
        //                    text.Text = text.Text.Replace(kvp.Key, kvp.Value);
        //                    Console.WriteLine($"{kvp.Key} => {kvp.Value}");
        //                }
        //            }
        //        }
        //        wordDoc.MainDocumentPart.Document.Save(); // Lưu thay đổi
        //    }
        //}
        private async void guna2Button2_luu_Click(object sender, EventArgs e)
        {
            if (!ValidateFormInputsV2(out string error))
            {
                ShowMessage(error, "Lỗi nhập liệu", Guna.UI2.WinForms.MessageDialogIcon.Warning);
                return;
            }

            try
            {
                string fileId = "1PFcPT9uq3XPlD9KrXiM27-wU4idW3Y9X";
                string tempPath = Path.Combine(Path.GetTempPath(), "Tangluong.docx");

                // Kiểm tra và xóa file temp nếu tồn tại
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }

                await DownloadFileFromGoogleDriveAsync(fileId, tempPath);

                // Kiểm tra file đã tải về có tồn tại không
                if (!File.Exists(tempPath))
                {
                    throw new Exception("Không thể tải file mẫu từ Google Drive");
                }

                var replacements = new Dictionary<string, string>
        {
            { "{{tenlanhdao}}", Lanhdao },
            { "{bophan}", Phongban },
            { "{luonght}", Salary_current.ToString("N0") + " VND"},
            { "{ten}", HoTen },
            { "{chucvu}", Vitri },
            { "{luongmm}", Salary_expected.ToString("N0") + " VND" },
            { "{day}", TuNgay },
            { "{lydo}", LyDo }
        };

                // Thêm log các giá trị sẽ thay thế
                Console.WriteLine("Các giá trị sẽ thay thế:");
                foreach (var kvp in replacements)
                {
                    Console.WriteLine($"{kvp.Key}: {kvp.Value}");
                }

                ReplacePlaceholders(tempPath, replacements);

                string defaultFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
                string suggestedPath = GetUniqueFileName(defaultFolder, "DonDeXuatTangLuong", ".docx");

                SaveFileDialog dialog = new SaveFileDialog
                {
                    Filter = "Word Documents (*.docx)|*.docx",
                    FileName = Path.GetFileName(suggestedPath),
                    InitialDirectory = defaultFolder
                };

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    File.Copy(tempPath, dialog.FileName, true);
                    ShowMessage("Lưu thành công!", "Thông báo", Guna.UI2.WinForms.MessageDialogIcon.Information);
                    ResetData();
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Lỗi khi tạo đơn: {ex.Message}", "Lỗi", Guna.UI2.WinForms.MessageDialogIcon.Error);
                Console.WriteLine($"Chi tiết lỗi: {ex.ToString()}");
            }
        }
        private string GetUniqueFileName(string folderPath, string baseName, string extension, int maxTries = 50)
        {
            for (int i = 1; i <= maxTries; i++)
            {
                string fileName = $"{baseName}_{i}{extension}";
                string fullPath = Path.Combine(folderPath, fileName);

                if (!File.Exists(fullPath))
                    return fullPath;
            }

            throw new IOException("Không thể tạo file mới. Đã đạt giới hạn số lần thử.");
        }
        //public static void ReplacePlaceholders(string filePath, Dictionary<string, string> replacements)
        //{
        //    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
        //    {
        //        var body = wordDoc.MainDocumentPart.Document.Body;

        //        foreach (var text in body.Descendants<Text>())
        //        {
        //            foreach (var kvp in replacements)
        //            {
        //                if (text.Text.Contains(kvp.Key))
        //                {
        //                    text.Text = text.Text.Replace(kvp.Key, kvp.Value);
        //                }
        //            }
        //        }

        //        wordDoc.MainDocumentPart.Document.Save();
        //    }
        //}
        private void ShowMessage(string message, string title, Guna.UI2.WinForms.MessageDialogIcon icon)
        {
            var mainForm = this.FindForm() as Home1;

            if (mainForm != null)
            {
                mainForm.ShowMessage(message, title, icon);
            }
            else
            {
                // Nếu không phải là Home1 hoặc không tìm thấy form
                MessageBox.Show(message, title, MessageBoxButtons.OK, GetMessageBoxIcon(icon));
            }
        }


        private MessageBoxIcon GetMessageBoxIcon(Guna.UI2.WinForms.MessageDialogIcon icon)
        {
            switch (icon)
            {
                case Guna.UI2.WinForms.MessageDialogIcon.Error:
                    return MessageBoxIcon.Error;
                case Guna.UI2.WinForms.MessageDialogIcon.Information:
                    return MessageBoxIcon.Information;
                case Guna.UI2.WinForms.MessageDialogIcon.Warning:
                    return MessageBoxIcon.Warning;
                case Guna.UI2.WinForms.MessageDialogIcon.Question:
                    return MessageBoxIcon.Question;
                default:
                    return MessageBoxIcon.None;
            }
        }


        private bool ValidateFormInputsV2(out string errorMessage)
        {
            StringBuilder sb = new StringBuilder();
            Regex onlyLettersRegex = new Regex(@"^[\p{L} ]+$"); // Cho phép chữ cái Unicode và khoảng trắng

            if (string.IsNullOrWhiteSpace(HoTen))
                sb.AppendLine("⚠️ Vui lòng nhập Họ tên.");
            else if (!onlyLettersRegex.IsMatch(HoTen))
                sb.AppendLine("⚠️ Họ tên chỉ được chứa chữ cái và khoảng trắng.");

            if (string.IsNullOrWhiteSpace(Phongban))
                sb.AppendLine("⚠️ Vui lòng nhập Phòng ban.");
            else if (!onlyLettersRegex.IsMatch(Phongban))
                sb.AppendLine("⚠️ Phòng ban chỉ được chứa chữ cái và khoảng trắng.");

            if (string.IsNullOrWhiteSpace(Vitri))
                sb.AppendLine("⚠️ Vui lòng nhập Vị trí công việc.");
            else if (!onlyLettersRegex.IsMatch(Vitri))
                sb.AppendLine("⚠️ Vị trí công việc chỉ được chứa chữ cái và khoảng trắng.");

            if (string.IsNullOrWhiteSpace(Lanhdao))
                sb.AppendLine("⚠️ Vui lòng nhập tên Lãnh đạo.");
            else if (!onlyLettersRegex.IsMatch(Lanhdao))
                sb.AppendLine("⚠️ Tên Lãnh đạo chỉ được chứa chữ cái và khoảng trắng.");

            if (string.IsNullOrWhiteSpace(LyDo))
                sb.AppendLine("⚠️ Vui lòng nhập Lý do.");
            string salaryCurrentStr = guna2NumericUpDown1_hientai.Text.Trim();
            string salaryExpectedStr = guna2NumericUpDown2_mongmuon.Text.Trim();
            if (!decimal.TryParse(salaryCurrentStr, out decimal Salary_current) || Salary_current <= 0)
            {
                sb.AppendLine("⚠️ Lương hiện tại phải là số hợp lệ và lớn hơn 0.");
            }
            else if (salaryCurrentStr.StartsWith("0") && salaryCurrentStr.Length > 1)
            {
                sb.AppendLine("⚠️ Lương hiện tại không được có số 0 ở đầu.");
            }

            if (!decimal.TryParse(salaryExpectedStr, out decimal Salary_expected) || Salary_expected <= 0)
            {
                sb.AppendLine("⚠️ Lương mong muốn phải là số hợp lệ và lớn hơn 0.");
            }
            else if (salaryExpectedStr.StartsWith("0") && salaryExpectedStr.Length > 1)
            {
                sb.AppendLine("⚠️ Lương mong muốn không được có số 0 ở đầu.");
            }

            if (guna2DateTimePicker1.Value.Date < DateTime.Today)
                sb.AppendLine("⚠️ Ngày bắt đầu không được nhỏ hơn ngày hôm nay.");

            errorMessage = sb.ToString();
            return string.IsNullOrEmpty(errorMessage);
        }

        private void DeXuatTangLuongControl_Load(object sender, EventArgs e)
        {
            CultureInfo viCulture = new CultureInfo("vi-VN");

            // Đặt mặc định cho toàn app (nên dùng)
            CultureInfo.DefaultThreadCurrentCulture = viCulture;
            CultureInfo.DefaultThreadCurrentUICulture = viCulture;

            guna2DateTimePicker1.Format = DateTimePickerFormat.Custom;
            guna2DateTimePicker1.CustomFormat = "'Ngày' dd 'tháng' MM 'năm' yyyy";
        }

        

        public void ResetData()
        {
            // Xóa nội dung các TextBox
            txt_lanhdao.Text = string.Empty;
            txt_name.Text = string.Empty;
            txt_bophan.Text = string.Empty;
            txt_chucvu.Text = string.Empty;

            // Xóa nội dung RichTextBox
            richTextBox1_lydo.Text = string.Empty;

            // Reset giá trị lương về 0 hoặc giá trị mặc định tùy bạn
            guna2NumericUpDown1_hientai.Value = 0;
            guna2NumericUpDown2_mongmuon.Value = 0;

            // Reset ngày về ngày hiện tại
            guna2DateTimePicker1.Value = DateTime.Now;
        }

       
            private void guna2Button1_huy_Click(object sender, EventArgs e)
            {
                // Tìm form chính
                var mainForm = this.FindForm() as Home1;

                // Tạo dialog xác nhận
                Guna.UI2.WinForms.Guna2MessageDialog dialog = new Guna.UI2.WinForms.Guna2MessageDialog
                {
                    Parent = mainForm, // đảm bảo dialog nằm giữa form cha
                    Buttons = Guna.UI2.WinForms.MessageDialogButtons.YesNo,
                    Caption = "Xác nhận",
                    Text = "Bạn có chắc muốn hủy thao tác không?",
                    Icon = Guna.UI2.WinForms.MessageDialogIcon.Question,
                    Style = Guna.UI2.WinForms.MessageDialogStyle.Light
                };

                // Hiển thị dialog và lấy kết quả
                var result = dialog.Show();

                if (result == DialogResult.Yes)
                {
                    // Gọi Reset nếu người dùng đồng ý hủy
                    ResetData();
                }
            }

      
    }

    }


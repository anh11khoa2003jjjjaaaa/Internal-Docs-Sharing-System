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
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Guna.UI2.WinForms;

namespace ChiaseNoiBo
{
    public partial class DonXinNghiPhepControl : UserControl
    {
        

        public DonXinNghiPhepControl()
        {
            InitializeComponent();
           
            
        }

        public string Lanhdao => txt_lanhdao.Text;
        public string HoTen => txt_name.Text;
        public string Phongban => txt_bophan.Text;
        public string Vitri => txt_chucvu.Text;
        public string LyDo => richTextBox1.Text;
        public string NguoiThayThe => txt_nguoithaythe.Text;
        public string BoPhanThayThe => txt_bophanthaythe.Text;
        public string TuNgay => guna2DateTime_batdau.Value.ToShortDateString();
        public string DenNgay => guna2DateTimePicker_ketthuc.Value.ToShortDateString();
       
        private void label1_Click(object sender, EventArgs e)
        {
            // Xử lý sự kiện click nếu cần
        }

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
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;

                foreach (var text in body.Descendants<Text>())
                {
                    foreach (var kvp in replacements)
                    {
                        if (text.Text.Contains(kvp.Key))
                        {
                            text.Text = text.Text.Replace(kvp.Key, kvp.Value);
                        }
                    }
                }

                wordDoc.MainDocumentPart.Document.Save();
            }
        }

        private async void guna2Button2_luu_Click(object sender, EventArgs e)
        {
            if (!ValidateFormInputs(out string errorMessage))
            {
                ShowMessage(errorMessage, "Thông báo lỗi", Guna.UI2.WinForms.MessageDialogIcon.Warning);
                return;
            }
            try
            {
                string fileId = "1I728PpgBMMdBiTBO5hWDTyEIea3RiWyr";
                string tempPath = Path.Combine(Path.GetTempPath(), "DonXinNghiPhep.docx");

                await DownloadFileFromGoogleDriveAsync(fileId, tempPath);

                // Tạo dictionary chứa các giá trị cần thay thế
                var replacements = new Dictionary<string, string>
                {
                    { "{{name}}", HoTen },
                    { "{{department}}", Phongban },
                    { "{{position}}", Vitri },
                    { "{{from_date}}", TuNgay },
                    { "{{to_date}}", DenNgay },
                    { "{{reason}}", LyDo },
                    { "{{recipient}}", Lanhdao },
                    { "{{handover_person}}", NguoiThayThe },
                    { "{{handover_department}}", BoPhanThayThe }
                
                };

                // Thay thế nội dung trong file
                ReplacePlaceholders(tempPath, replacements);

                // Lấy thư mục Downloads của người dùng hoặc Temp tùy ý
                string defaultFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),"Downloads");
                // hoặc Downloads

                // Gọi hàm để tạo tên file tự động
                string suggestedPath = GetUniqueFileName(defaultFolder, "DonXinNghiPhep", ".docx");

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
                ShowMessage("Lỗi: " + ex.Message, "Lỗi", Guna.UI2.WinForms.MessageDialogIcon.Error);
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

        
        private void ShowMessage(string message, string title, Guna.UI2.WinForms.MessageDialogIcon icon)
        {
            var mainForm = this.FindForm() as Home1;

            if (mainForm != null)
            {
                mainForm.ShowMessage(message, title, icon);
            }
            else
            {
                // Tạo Guna2MessageDialog tạm nếu không có form chính
                Guna.UI2.WinForms.Guna2MessageDialog dialog = new Guna.UI2.WinForms.Guna2MessageDialog();
                dialog.Buttons = Guna.UI2.WinForms.MessageDialogButtons.OK;
                dialog.Icon = icon;
                dialog.Style = Guna.UI2.WinForms.MessageDialogStyle.Light;
                dialog.Caption = title;
                dialog.Text = message;

                // Đảm bảo dialog nằm giữa màn hình
                
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
        public void ResetData()
        {
            // Xóa dữ liệu trong các TextBox
            txt_lanhdao.Text = string.Empty;
            txt_name.Text = string.Empty;
            txt_bophan.Text = string.Empty;
            txt_chucvu.Text = string.Empty;
            txt_nguoithaythe.Text = string.Empty;
            txt_bophanthaythe.Text = string.Empty;

            // Xóa nội dung RichTextBox
            richTextBox1.Text = string.Empty;

            // Reset ngày về ngày hiện tại
            guna2DateTime_batdau.Value = DateTime.Now;
            guna2DateTimePicker_ketthuc.Value = DateTime.Now;
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

        private bool ValidateFormInputs(out string errorMessage)
        {
            StringBuilder sb = new StringBuilder();

            if (string.IsNullOrWhiteSpace(HoTen) || HoTen.All(char.IsDigit))
                sb.AppendLine("⚠️ Họ tên không hợp lệ (không được để trống hoặc chỉ gồm số).");

            if (string.IsNullOrWhiteSpace(Phongban) || Phongban.All(char.IsDigit))
                sb.AppendLine("⚠️ Phòng ban không hợp lệ (không được để trống hoặc chỉ gồm số).");

            if (string.IsNullOrWhiteSpace(Vitri) || Vitri.All(char.IsDigit))
                sb.AppendLine("⚠️ Vị trí công việc không hợp lệ (không được để trống hoặc chỉ gồm số).");

            if (string.IsNullOrWhiteSpace(Lanhdao) || Lanhdao.All(char.IsDigit))
                sb.AppendLine("⚠️ Tên lãnh đạo không hợp lệ (không được để trống hoặc chỉ gồm số).");

            if (string.IsNullOrWhiteSpace(LyDo) || LyDo.All(char.IsDigit))
                sb.AppendLine("⚠️ Lý do xin nghỉ không hợp lệ (không được để trống hoặc chỉ gồm số).");

            if (string.IsNullOrWhiteSpace(NguoiThayThe) || NguoiThayThe.All(char.IsDigit))
                sb.AppendLine("⚠️ Tên người thay thế không hợp lệ (không được để trống hoặc chỉ gồm số).");

            if (string.IsNullOrWhiteSpace(BoPhanThayThe) || BoPhanThayThe.All(char.IsDigit))
                sb.AppendLine("⚠️ Bộ phận người thay thế không hợp lệ (không được để trống hoặc chỉ gồm số).");

        var tuNgay = guna2DateTime_batdau.Value.Date;
            var denNgay = guna2DateTimePicker_ketthuc.Value.Date;

            if (denNgay < tuNgay)
            {
                sb.AppendLine("❌ Ngày kết thúc không được trước ngày bắt đầu nghỉ!");
            }

            errorMessage = sb.ToString();
            return string.IsNullOrEmpty(errorMessage);
        }

        private void DonXinNghiPhepControl_Load(object sender, EventArgs e)
        {
            CultureInfo viCulture = new CultureInfo("vi-VN");

            // Đặt mặc định cho toàn app (nên dùng)
            CultureInfo.DefaultThreadCurrentCulture = viCulture;
            CultureInfo.DefaultThreadCurrentUICulture = viCulture;

            guna2DateTime_batdau.Format = DateTimePickerFormat.Custom;
            guna2DateTime_batdau.CustomFormat = "'Ngày' dd 'tháng' MM 'năm' yyyy";
            guna2DateTimePicker_ketthuc.Format = DateTimePickerFormat.Custom;
            guna2DateTimePicker_ketthuc.CustomFormat = "'Ngày' dd 'tháng' MM 'năm' yyyy";

        }
    }
}
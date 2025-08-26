
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Guna.UI2.WinForms;
using System.Security.Cryptography;
using System.Linq;

namespace ChiaseNoiBo
{
    public partial class Register : Form
    {
        private const string EXCEL_FILENAME = "DanhSachTaiKhoan.xlsx";
        private readonly Color ERROR_COLOR = Color.Red;
        private readonly Color PLACEHOLDER_COLOR = Color.Gray;

        public Register()
        {
            InitializeComponent();
            SetupPlaceholders();
            txt_password.UseSystemPasswordChar = true;
            confirm_password.UseSystemPasswordChar = true;
            guna2PictureBox1.Image = Properties.Resources.hidden;
            guna2PictureBox2.Image = Properties.Resources.hidden;
        }

        private void SetupPlaceholders()
        {
            txt_name.PlaceholderText = "Nhập họ và tên";
            txt_email.PlaceholderText = "Nhập email";
            txt_password.PlaceholderText = "Nhập mật khẩu";
            confirm_password.PlaceholderText = "Nhập lại mật khẩu";
            txt_name.PlaceholderForeColor = PLACEHOLDER_COLOR;
            txt_email.PlaceholderForeColor = PLACEHOLDER_COLOR;
            txt_password.PlaceholderForeColor = PLACEHOLDER_COLOR;
            confirm_password.PlaceholderForeColor = PLACEHOLDER_COLOR;
        }

        private async Task UploadExcelFileAsync(string filePath)
        {
            var service = GoogleDriveHelper.GetDriveService();

            var listRequest = service.Files.List();
            listRequest.Q = $"name='{EXCEL_FILENAME}' and '{GoogleDriveHelper.FolderId}' in parents and trashed=false";
            listRequest.Fields = "files(id)";
            var result = await listRequest.ExecuteAsync();

            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                if (result.Files.Count > 0)
                {
                    var updateRequest = service.Files.Update(null, result.Files[0].Id, stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                    await updateRequest.UploadAsync();
                }
                else
                {
                    var fileMetadata = new Google.Apis.Drive.v3.Data.File()
                    {
                        Name = EXCEL_FILENAME,
                        Parents = new List<string> { GoogleDriveHelper.FolderId }
                    };
                    var createRequest = service.Files.Create(fileMetadata, stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                    await createRequest.UploadAsync();
                }
            }
        }

        private async Task SyncToGoogleDrive(string localFilePath)
        {
            await UploadExcelFileAsync(localFilePath);
        }

        private void ResetForm()
        {
            txt_name.Text = "";
            txt_email.Text = "";
            txt_password.Text = "";
            confirm_password.Text = "";
            SetupPlaceholders();
        }

        private bool isValidate()
        {
            bool isValid = true;

            if (string.IsNullOrWhiteSpace(txt_name.Text))
            {
                SetError(txt_name, "Họ tên không được để trống!");
                isValid = false;
            }

            string email = txt_email.Text.Trim();
            if (string.IsNullOrWhiteSpace(email))
            {
                SetError(txt_email, "Email không được để trống!");
                isValid = false;
            }
            else
            {
                try
                {
                    var addr = new MailAddress(email);
                }
                catch
                {
                    SetError(txt_email, "Email không hợp lệ!");
                    isValid = false;
                }
            }

            string password = txt_password.Text.Trim();
            if (string.IsNullOrWhiteSpace(password) || string.IsNullOrWhiteSpace(confirm_password.Text))
            {
                SetError(txt_password, "Mật khẩu không được để trống!");
                SetError(confirm_password, "Mật khẩu không được để trống!");
                isValid = false;
            }
            else if (password.EndsWith("_"))
            {
                SetError(txt_password, "Mật khẩu không được kết thúc bằng dấu gạch dưới (_)");
                isValid = false;
            }

            if (confirm_password.Text.Trim() != password)
            {
                SetError(confirm_password, "Mật khẩu không đúng");
                isValid = false;
            }

            return isValid;
        }

        private void SetError(Guna2TextBox textBox, string errorMessage)
        {
            textBox.Text = "";
            textBox.PlaceholderText = errorMessage;
            textBox.PlaceholderForeColor = Color.Red;
        }

        private async Task<bool> CheckEmailExistsInExcelAsync(string emailToCheck)
        {
            var service = GoogleDriveHelper.GetDriveService();
            var listRequest = service.Files.List();
            listRequest.Q = $"name='{EXCEL_FILENAME}' and '{GoogleDriveHelper.FolderId}' in parents and trashed=false";
            listRequest.Fields = "files(id)";
            var result = await listRequest.ExecuteAsync();

            if (result.Files == null || result.Files.Count == 0)
                throw new Exception("Không tìm thấy file DanhSachTaiKhoan.xlsx trên Google Drive");

            string fileId = result.Files[0].Id;
            string localTempPath = Path.Combine(Path.GetTempPath(), EXCEL_FILENAME);

            using (var stream = new FileStream(localTempPath, FileMode.Create, FileAccess.Write))
            {
                await service.Files.Get(fileId).DownloadAsync(stream);
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(localTempPath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null) return false;

                int rowCount = worksheet.Dimension?.End.Row ?? 0;

                for (int row = 2; row <= rowCount; row++)
                {
                    var emailInCell = worksheet.Cells[row, 3]?.Value?.ToString();
                    if (!string.IsNullOrEmpty(emailInCell) && emailInCell.Equals(emailToCheck, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        public string HashPasswordSHA256(string password)
        {
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] inputBytes = Encoding.UTF8.GetBytes(password);
                byte[] hashBytes = sha256.ComputeHash(inputBytes);

                StringBuilder sb = new StringBuilder();
                foreach (byte b in hashBytes)
                    sb.Append(b.ToString("x2"));

                return sb.ToString();
            }
        }

        private bool IsStrongPassword(string password)
        {
            var regex = new System.Text.RegularExpressions.Regex(@"^(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*[@$!%*?&])[A-Za-z\d@$!%*?&]{8,}$");
            return regex.IsMatch(password);
        }

        private async Task<string> DownloadExcelFileAsync()
        {
            var service = GoogleDriveHelper.GetDriveService();
            string tempFilePath = Path.Combine(Path.GetTempPath(), EXCEL_FILENAME);

            var listRequest = service.Files.List();
            listRequest.Q = $"name='{EXCEL_FILENAME}' and '{GoogleDriveHelper.FolderId}' in parents and trashed=false";
            listRequest.Fields = "files(id)";
            var result = await listRequest.ExecuteAsync();

            if (result.Files.Count == 0)
            {
                CreateNewExcelFile(tempFilePath);
            }
            else
            {
                using (var stream = new FileStream(tempFilePath, FileMode.Create))
                {
                    await service.Files.Get(result.Files[0].Id).DownloadAsync(stream);
                }
            }

            return tempFilePath;
        }

        private void CreateNewExcelFile(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.Add("Accounts");
                worksheet.Cells["A1"].Value = "STT";
                worksheet.Cells["B1"].Value = "Họ và tên";
                worksheet.Cells["C1"].Value = "Email";
                worksheet.Cells["D1"].Value = "Password";
                worksheet.Cells["E1"].Value = "Xác thực";
                worksheet.Cells["F1"].Value = "Vai trò";
                package.Save();
            }
        }

        private bool AddUserToExcel(string filePath, string name, string email, string hashedPassword)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                for (int row = 2; row <= (worksheet.Dimension?.End.Row ?? 1); row++)
                {
                    if (worksheet.Cells[row, 3].Text.Equals(email, StringComparison.OrdinalIgnoreCase))
                        return false;
                }

                int newRow = (worksheet.Dimension?.End.Row ?? 1) + 1;
                worksheet.Cells[newRow, 1].Value = newRow - 1;
                worksheet.Cells[newRow, 2].Value = name;
                worksheet.Cells[newRow, 3].Value = email;
                worksheet.Cells[newRow, 4].Value = hashedPassword;
                worksheet.Cells[newRow, 5].Value = "Chưa hoạt động";
                worksheet.Cells[newRow, 6].Value = "User";

                package.Save();
                return true;
            }
        }
       

        private async void guna2Button1_Click(object sender, EventArgs e)
        {
            if (!isValidate()) return;

            string name = txt_name.Text.Trim();
            string email = txt_email.Text.Trim();
            string password = txt_password.Text.Trim();

            if (!IsStrongPassword(password))
            {
                ShowMessage("Mật khẩu phải có ít nhất 8 ký tự, gồm chữ hoa, chữ thường, số và ký tự đặc biệt.", "Cảnh báo", MessageDialogIcon.Warning);
                return;
            }

            string hashedPassword = HashPasswordSHA256(password + "_phcn");

            try
            {
                if (await CheckEmailExistsInExcelAsync(email))
                {
                    ShowMessage("Email đã tồn tại trong hệ thống!", "Cảnh báo", MessageDialogIcon.Warning);
                    return;
                }

                string tempFilePath = await DownloadExcelFileAsync();
                bool isNewRecordAdded = AddUserToExcel(tempFilePath, name, email, hashedPassword);

                if (isNewRecordAdded)
                {
                    await SyncToGoogleDrive(tempFilePath);
                    ShowMessage("Đăng ký thành công!Vui lòng chờ Admin duyệt", "Thông báo", MessageDialogIcon.Information);
                    ResetForm();
                    new Login().Show();
                    this.Hide();
                }
                else
                {
                    ShowMessage("Đăng ký thất bại! Vui lòng nhập lại", "Thông báo lỗi", MessageDialogIcon.Error);
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Lỗi: {ex.Message}", "Lỗi", MessageDialogIcon.Error);
            }
        }

        private void ShowMessage(string message, string title, Guna.UI2.WinForms.MessageDialogIcon icon)
        {
            var dialog = new Guna.UI2.WinForms.Guna2MessageDialog
            {
                Buttons = Guna.UI2.WinForms.MessageDialogButtons.OK,
                Icon = icon,
                Style = Guna.UI2.WinForms.MessageDialogStyle.Dark,
                Caption = title,
                Text = message,
                Parent = this
            };
            dialog.Show();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            new Login().Show();
            this.Hide();
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void guna2PictureBox1_Click(object sender, EventArgs e)
        {
            txt_password.UseSystemPasswordChar = !txt_password.UseSystemPasswordChar;
            guna2PictureBox1.Image = txt_password.UseSystemPasswordChar
                ? Properties.Resources.hidden
                : Properties.Resources.eye;
        }

        private void guna2PictureBox2_Click(object sender, EventArgs e)
        {
            confirm_password.UseSystemPasswordChar = !confirm_password.UseSystemPasswordChar;
            guna2PictureBox2.Image = confirm_password.UseSystemPasswordChar
                ? Properties.Resources.hidden
                : Properties.Resources.eye;
        }
    }
}
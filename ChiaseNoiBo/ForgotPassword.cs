using MimeKit;
using System;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MailKit.Net.Smtp;
using MailKit.Security;
using System.Drawing;
using Google.Apis.Drive.v3;
using OfficeOpenXml;
using System.IO;
using System.Threading.Tasks;
using Guna.UI2.WinForms;

namespace ChiaseNoiBo
{
    public partial class ForgotPassword : Form
    {
        public ForgotPassword()
        {
            InitializeComponent();
        }

        

        //Hàm xác nhận gửi mail
        private async void guna2Button1_Click(object sender, EventArgs e)
        {
            // Kiểm tra và xác thực email
            string email = txt_email.Text.Trim();
            if (!IsValidEmail(email))
            {
                ShowMessage("Vui lòng nhập một địa chỉ email hợp lệ.", "Lỗi", MessageDialogIcon.Error);
                return;
            }

            // Tạo mật khẩu mới ngẫu nhiên và băm nó
            string newPassword = GenerateRandomPassword(8);
            string hashedPassword = HashSHA256(newPassword);

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                // Gửi email với mật khẩu mới
                SendResetPasswordEmailWithMailKit(email, newPassword);

                // Cập nhật mật khẩu mới vào file trên Google Drive
                bool isUpdated = await UpdatePasswordOnGoogleDrive(email, hashedPassword);
                if (isUpdated)
                {
                    ShowMessage("Mật khẩu mới đã được cập nhật và gửi qua email!", "Thông báo", MessageDialogIcon.Information);
                    Login login = new Login();
                    login.Show();
                    this.Hide();
                }
                else
                {
                    ShowMessage("Không tìm thấy tài khoản với email này trong danh sách.", "Lỗi", MessageDialogIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Đã xảy ra lỗi trong quá trình xử lý: {ex.Message}", "Lỗi", MessageDialogIcon.Error);
            }
        }

        private bool IsValidEmail(string email)
        {
            // Kiểm tra rỗng và định dạng email hợp lệ
            if (string.IsNullOrEmpty(email))
            {
                txt_email.PlaceholderText = "Email không được để trống!";
                txt_email.PlaceholderForeColor = Color.Red;
                txt_email.Text = ""; // Xoá nội dung để hiển thị placeholder
                return false;
            }

            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return true;
            }
            catch
            {
                txt_email.Text = "";
                txt_email.PlaceholderText = "Email không hợp lệ!";
                txt_email.PlaceholderForeColor = Color.Red;
                return false;
            }
        }

        private string GenerateRandomPassword(int length)
        {
            var random = new Random();
            const string chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
                .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        private string HashSHA256(string input)
        {
            using (var sha256 = System.Security.Cryptography.SHA256.Create())
            {
                byte[] bytes = Encoding.UTF8.GetBytes(input);
                byte[] hashBytes = sha256.ComputeHash(bytes);
                return BitConverter.ToString(hashBytes).Replace("-", "").ToLower();
            }
        }

        private void SendResetPasswordEmailWithMailKit(string toEmail, string newPassword)
        {
            try
            {
                var message = new MimeMessage();
                message.From.Add(new MailboxAddress("ChiaseNoiBo App", "huynhanhkhoa30042019@gmail.com"));
                message.To.Add(new MailboxAddress("User", toEmail));
                message.Subject = "Khôi phục mật khẩu - ChiaseNoiBo";

                message.Body = new TextPart("plain")
                {
                    Text = $"Xin chào,\n\nMật khẩu mới của bạn là: {newPassword}\nVui lòng đăng nhập và đổi lại mật khẩu.\n\nTrân trọng."
                };

                using (var client = new SmtpClient())
                {
                    client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                    client.Authenticate("huynhanhkhoa30042019@gmail.com", "pprn eagw zjwa lzwq");
                    client.Send(message);
                    client.Disconnect(true);
                }
            }
            catch (Exception ex)
            {
               ShowMessage($"Gửi mail thất bại: {ex.Message}", "Lỗi", MessageDialogIcon.Error);
            }
        }

        private async Task<bool> UpdatePasswordOnGoogleDrive(string email, string newPassword)
        {
            var service = GoogleDriveHelper.GetDriveService();
            string fileName = "DanhSachTaiKhoan.xlsx";

            try
            {
                // Lấy ID của file trên Google Drive
                var fileId = await GetFileIdOnDrive(service, fileName);
                if (fileId == null)
                {
                    ShowMessage("Không tìm thấy file trên Google Drive", "Lỗi", MessageDialogIcon.Error);
                    return false;
                }

                using (var memoryStream = new MemoryStream())
                {
                    var request = service.Files.Get(fileId);
                    await request.DownloadAsync(memoryStream);

                    // Cập nhật mật khẩu trong bộ nhớ
                    memoryStream.Position = 0;
                    using (var package = new ExcelPackage(memoryStream))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        bool isUpdated = false;

                        for (int row = 2; row <= worksheet.Dimension?.End.Row; row++)
                        {
                            if (worksheet.Cells[row, 3]?.Value?.ToString() == email)
                            {
                                worksheet.Cells[row, 4].Value = newPassword;
                                isUpdated = true;
                                break;
                            }
                        }

                        if (isUpdated)
                        {
                            // Lưu lại thay đổi và cập nhật lên Google Drive
                            var updatedMemoryStream = new MemoryStream();
                            package.SaveAs(updatedMemoryStream);

                            updatedMemoryStream.Position = 0;
                            var updateRequest = service.Files.Update(
                                null, fileId, updatedMemoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                            await updateRequest.UploadAsync();

                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Cập nhật mật khẩu trên Google Drive thất bại: {ex.Message}", "Lỗi", MessageDialogIcon.Error);
            }

            return false;
        }

        private async Task<string> GetFileIdOnDrive(DriveService service, string fileName)
        {
            var listRequest = service.Files.List();
            listRequest.Q = $"name='{fileName}' and '{GoogleDriveHelper.FolderId}' in parents and trashed=false";
            listRequest.Fields = "files(id)";

            var fileList = await listRequest.ExecuteAsync();
            return fileList.Files.Any() ? fileList.Files[0].Id : null;
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            this.Close();
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
            Login login= new Login();
            login.Show();
            this.Hide();
        }
    }
  
}

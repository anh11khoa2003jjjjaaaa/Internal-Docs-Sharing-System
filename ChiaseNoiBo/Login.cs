using OfficeOpenXml;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ChiaseNoiBo
{
    public partial class Login : Form
    {


        public Login()
        {
            InitializeComponent();
            txt_password.UseSystemPasswordChar = true; // Ẩn pass ngay từ đầu
            guna2PictureBox1.Image = Properties.Resources.hidden; // Con mắt đóng
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            this.Close(); // Đóng form đăng nhập
        }

        private async void guna2Button1_Click(object sender, EventArgs e)
        {
            await PerformLogin();
        }

        private async void txt_password_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true; // Ngăn tiếng "bíp"
                await PerformLogin();
            }
        }
        public static string HashPasswordSHA256(string password)
        {
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] inputBytes = Encoding.UTF8.GetBytes(password);
                byte[] hashBytes = sha256.ComputeHash(inputBytes);

                // Convert bytes to hex string
                StringBuilder sb = new StringBuilder();
                foreach (byte b in hashBytes)
                    sb.Append(b.ToString("x2")); // "x2" to get hex format

                return sb.ToString();
            }
        }
        private async Task PerformLogin()
        {
            string email = txt_email.Text.Trim();
            string password = txt_password.Text.Trim();
            string hashedPassword = HashPasswordSHA256(password);

            var (username, role) = await AuthenticateUser(email, hashedPassword);

            if (username != null)
            {
                await AuthenticateUser(email, hashedPassword);
                ShowMessage("✅ Đăng nhập thành công!", "Thông báo", Guna.UI2.WinForms.MessageDialogIcon.Information);
                this.Hide();
                new Home1(username, role).Show();
            }
        
        }

        private async Task<string> FindFileDanhSachTaiKhoanfromGGdrive()
        {
            try
            {
                var service = GoogleDriveHelper.GetDriveService();
                var request = service.Files.List();
                request.Q = $"'{GoogleDriveHelper.FolderId}' in parents and name = 'DanhSachTaiKhoan.xlsx'";
                request.Fields = "files(id, name)";

                var result = await request.ExecuteAsync();
                if (result.Files.Count == 0) return null;

                var fileId = result.Files[0].Id;
                string localFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DanhSachTaiKhoan.xlsx");

                using (var stream = new MemoryStream())
                {
                    await service.Files.Get(fileId).DownloadAsync(stream);
                    File.WriteAllBytes(localFilePath, stream.ToArray());
                }
                return localFilePath;
            }
            catch (Exception ex)
            {
                ShowMessage($"Lỗi khi tải file từ Google Drive: {ex.Message}", "Lỗi", Guna.UI2.WinForms.MessageDialogIcon.Error);
                return null;
            }
        }

        private async Task<(string username, string role)> AuthenticateUser(string email, string password)
        {
            string excelPath = await FindFileDanhSachTaiKhoanfromGGdrive();
            if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
            {
                ShowMessage("Không tìm thấy dữ liệu tài khoản!", "Lỗi", Guna.UI2.WinForms.MessageDialogIcon.Error);
                return (null, null);
            }

            var (username, status, role) = CheckCredentialsAndStatusFromExcel(excelPath, email, password);

            if (username == null)
            {
                ShowMessage("❌ Tài khoản hoặc mật khẩu không chính xác!", "Lỗi", Guna.UI2.WinForms.MessageDialogIcon.Error);
                return (null, null);
            }

            if (status != "Hoạt động")
            {
                ShowMessage("⚠️ Tài khoản chưa được kích hoạt hoặc đã bị khóa!", "Cảnh báo", Guna.UI2.WinForms.MessageDialogIcon.Warning);
                return (null, null);
            }

            return (username, role);
        }

        //private async Task<(string username, string role)> AuthenticateUser(string email, string password)
        //{
        //    string excelPath = await FindFileDanhSachTaiKhoanfromGGdrive();
        //    if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath)) return (null, null);

        //    var (username, status, role) = CheckCredentialsAndStatusFromExcel(excelPath, email, password);

        //    if (status != "Hoạt động")
        //    {
        //        ShowMessage("Tài khoản chưa được kích hoạt hoặc đã bị khóa!", "Cảnh báo", Guna.UI2.WinForms.MessageDialogIcon.Warning);
        //        return (null, null);
        //    }

        //    return (username, role);
        //}

        private (string username, string status, string role) CheckCredentialsAndStatusFromExcel(string filePath, string email, string password)
        {
            try
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        string excelEmail = worksheet.Cells[row, 3].Text.Trim();
                        string username = worksheet.Cells[row, 2].Text.Trim();
                        string pass = worksheet.Cells[row, 4].Text.Trim();
                        string status = worksheet.Cells[row, 5].Text.Trim();
                        string role = worksheet.Cells[row, 6].Text.Trim(); // Giả sử cột 6 là Vai trò

                        if (email == excelEmail && password == pass)
                        {
                            return (username, status, role);
                        }
                    }
                }
                return (null, null, null);
            }
            catch (Exception ex)
            {
                ShowMessage($"Lỗi khi đọc file Excel: {ex.Message}", "Lỗi", Guna.UI2.WinForms.MessageDialogIcon.Error);
                return (null, null, null);
            }
        }



        private void ShowMessage(string message, string title, Guna.UI2.WinForms.MessageDialogIcon icon)
        {
            guna2MessageDialog1.Icon = icon;
            guna2MessageDialog1.Show(message, title);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Register register = new Register();
            register.Show();
            this.Hide();
        }

        private void guna2PictureBox1_Click(object sender, EventArgs e)
        {
            txt_password.UseSystemPasswordChar = !txt_password.UseSystemPasswordChar;
            guna2PictureBox1.Image = txt_password.UseSystemPasswordChar
                ? Properties.Resources.hidden
                : Properties.Resources.eye;
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            ForgotPassword formForgot= new ForgotPassword();
            formForgot.Show();
            this.Hide();
        }
    }
}
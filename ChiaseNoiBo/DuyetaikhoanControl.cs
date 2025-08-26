using Guna.UI2.WinForms;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using System.Drawing;
using MailKit.Security;
using MimeKit;
using MailKit.Net.Smtp;

namespace ChiaseNoiBo
{
    public partial class DuyetaikhoanControl : UserControl
    {
        private const string EXCEL_FILENAME = "DanhSachTaiKhoan.xlsx";

        public DuyetaikhoanControl()
        {
            InitializeComponent();
            InitDataGridView();
            _ = LoadAndDisplayExcelDataAsync();
        }

        private void InitDataGridView()
        {
            guna2DataGridView1.Columns.Clear();
            guna2DataGridView1.AllowUserToAddRows = false;
            guna2DataGridView1.RowHeadersVisible = false;
            guna2DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            guna2DataGridView1.Columns.Add("STT", "STT");
            guna2DataGridView1.Columns.Add("HoTen", "Họ và tên");
            guna2DataGridView1.Columns.Add("Email", "Email");
            guna2DataGridView1.Columns.Add("Password", "Password");
            guna2DataGridView1.Columns.Add("XacThuc", "Xác thực");
            guna2DataGridView1.Columns.Add("VaiTro", "Vai trò");

            var checkBoxCol = new DataGridViewCheckBoxColumn()
            {
                Name = "HanhDong",
                HeaderText = "Duyệt?",
                Width = 60
            };
            guna2DataGridView1.Columns.Add(checkBoxCol);
        }

        private async Task LoadAndDisplayExcelDataAsync()
        {
            try
            {
                string filePath = await DownloadExcelFileAsync();

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var ws = package.Workbook.Worksheets[0];
                    int rows = ws.Dimension?.End.Row ?? 1;

                    guna2DataGridView1.Rows.Clear();
                    for (int row = 2; row <= rows; row++)
                    {
                        guna2DataGridView1.Rows.Add(
                            ws.Cells[row, 1].Text,
                            ws.Cells[row, 2].Text,
                            ws.Cells[row, 3].Text,
                            ws.Cells[row, 4].Text,
                            ws.Cells[row, 5].Text,
                            ws.Cells[row, 6].Text,
                            ws.Cells[row, 5].Text == "Hoạt động"
                        );
                    }
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Lỗi khi tải dữ liệu: {ex.Message}", "Lỗi", MessageDialogButtons.OK, MessageDialogIcon.Error);
            }
        }

        private async Task<string> DownloadExcelFileAsync()
        {
            var service = GoogleDriveHelper.GetDriveService();
            string tempPath = Path.Combine(Path.GetTempPath(), EXCEL_FILENAME);

            var listRequest = service.Files.List();
            listRequest.Q = $"name='{EXCEL_FILENAME}' and '{GoogleDriveHelper.FolderId}' in parents and trashed=false";
            listRequest.Fields = "files(id)";
            var result = await listRequest.ExecuteAsync();

            if (result.Files.Count == 0)
            {
                CreateNewExcelFile(tempPath);
            }
            else
            {
                using (var stream = new FileStream(tempPath, FileMode.Create))
                {
                    await service.Files.Get(result.Files[0].Id).DownloadAsync(stream);
                }
            }

            return tempPath;
        }

        private void CreateNewExcelFile(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var ws = package.Workbook.Worksheets.Add("Accounts");
                ws.Cells["A1"].Value = "STT";
                ws.Cells["B1"].Value = "Họ và tên";
                ws.Cells["C1"].Value = "Email";
                ws.Cells["D1"].Value = "Password";
                ws.Cells["E1"].Value = "Xác thực";
                ws.Cells["F1"].Value = "Vai trò";
                package.Save();
            }
        }

        private async Task UpdateExcelAndSync(string email, string newStatus)
        {
            try
            {
                string filePath = await DownloadExcelFileAsync();

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var ws = package.Workbook.Worksheets[0];
                    int rows = ws.Dimension?.End.Row ?? 1;

                    for (int i = 2; i <= rows; i++)
                    {
                        if (ws.Cells[i, 3].Text.Equals(email, StringComparison.OrdinalIgnoreCase))
                        {
                            ws.Cells[i, 5].Value = newStatus;
                            break;
                        }
                    }

                    package.Save();
                }

                await UploadExcelFileAsync(filePath);
            }
            catch (Exception ex)
            {
                ShowMessage($"Lỗi khi cập nhật dữ liệu: {ex.Message}", "Lỗi", MessageDialogButtons.OK, MessageDialogIcon.Error);
            }
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

        private async void guna2DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (guna2DataGridView1.Columns[e.ColumnIndex].Name == "HanhDong" && e.RowIndex >= 0)
            {
               


                var currentRow = guna2DataGridView1.Rows[e.RowIndex];
                string email = currentRow.Cells["Email"].Value?.ToString();
                if (string.IsNullOrEmpty(email)) return;

               
                string hanhdong = currentRow.Cells["XacThuc"].Value?.ToString();

                if ( hanhdong == "Chưa hoạt động")
                {
                    // DUYỆT
                    var dialog = new Guna2MessageDialog()
                    {
                        Text = $"Bạn có chắc chắn muốn duyệt tài khoản {email}?",
                        Caption = "Xác nhận duyệt",
                        Buttons = MessageDialogButtons.YesNo,
                        Icon = MessageDialogIcon.Question,
                        Style = MessageDialogStyle.Light,
                        Parent=this.FindForm()
                    };

                    var result = dialog.Show();

                    if (result == DialogResult.Yes)
                    {
                        currentRow.Cells["XacThuc"].Value = "Hoạt động";
                        currentRow.Cells["HanhDong"].Value = true;
                        await UpdateExcelAndSync(email, "Hoạt động");
                        ShowMessage("Đã cập nhật thành công!", "Thông báo", MessageDialogButtons.OK, MessageDialogIcon.Information);
                    }
                    else
                    {
                        currentRow.Cells["HanhDong"].Value = false; // Khôi phục lại nếu người dùng chọn No
                    }
                }
                else if (hanhdong == "Hoạt động")
                {
                    // HỦY DUYỆT
                    var dialog = new Guna2MessageDialog()
                    {
                        Text = $"Bạn có chắc chắn muốn hủy duyệt tài khoản {email}?",
                        Caption = "Xác nhận hủy duyệt",
                        Buttons = MessageDialogButtons.YesNo,
                        Icon = MessageDialogIcon.Warning,
                        Style = MessageDialogStyle.Light,
                        Parent = this.FindForm()
                    };

                    var result = dialog.Show();

                    if (result == DialogResult.Yes)
                    {
                        currentRow.Cells["XacThuc"].Value = "Đã hủy";
                        currentRow.Cells["HanhDong"].Value = false;
                        await UpdateExcelAndSync(email, "Đã hủy");
                        ShowMessage("Đã cập nhật thành công!", "Thông báo", MessageDialogButtons.OK, MessageDialogIcon.Information);
                        SendAccountDeactivationEmailWithMailKit(email);
                    }
                    else
                    {
                        currentRow.Cells["HanhDong"].Value = true; // Khôi phục lại nếu người dùng chọn No
                    }
                }
            }
        }

        private void ShowMessage(string message, string title, MessageDialogButtons buttons, MessageDialogIcon icon)
        {
            var dialog = new Guna2MessageDialog()
            {
                Text = message,
                Caption = title,
                Buttons = buttons,
                Icon = icon,
                Style = MessageDialogStyle.Light,
                Parent = this.FindForm()
            };

            dialog.Show();
        }

        private void guna2DataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            var row = guna2DataGridView1.Rows[e.RowIndex];
            if (row.Cells["XacThuc"].Value?.ToString() == "Đã hủy")
            {
                row.DefaultCellStyle.ForeColor = Color.Gray;
                row.DefaultCellStyle.BackColor = Color.LightGray;
                row.ReadOnly = true; // Không cho sửa
            }
        }

        private void SendAccountDeactivationEmailWithMailKit(string toEmail)
        {
            try
            {
                var message = new MimeMessage();
                message.From.Add(new MailboxAddress("ChiaseNoiBo App", "huynhanhkhoa30042019@gmail.com"));
                message.To.Add(new MailboxAddress("User", toEmail));
                message.Subject = "Thông báo hủy tài khoản - ChiaseNoiBo";

                message.Body = new TextPart("plain")
                {
                    Text =
     "Xin chào,\n\n" +
     "Chúng tôi xin thông báo rằng tài khoản của bạn đã bị hủy.\n\n" +
     "Lý do: Bạn không còn là thành viên của cơ quan chúng tôi.\n\n" +
     "Nếu có bất kỳ thắc mắc nào, vui lòng liên hệ với chúng tôi.\n\n" +
     "Trân trọng.\n\n" +
     "---\n\n" +
     "Hello,\n\n" +
     "We would like to inform you that your account has been deactivated.\n\n" +
     "Reason: You are no longer a member of our organization.\n\n" +
     "If you have any questions, please feel free to contact us.\n\n" +
     "Best regards."
                };


                using (var client = new SmtpClient())
                {
                    // Kết nối đến SMTP server của Gmail và gửi email
                    client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                    client.Authenticate("huynhanhkhoa30042019@gmail.com", "pprn eagw zjwa lzwq");
                    client.Send(message);
                    client.Disconnect(true);
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Gửi mail thất bại: {ex.Message}", "Lỗi",MessageDialogButtons.OK, MessageDialogIcon.Error);
            }
        }

        private void guna2DataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (guna2DataGridView1.IsCurrentCellDirty)
            {
                guna2DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }
    }
}
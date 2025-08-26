using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Guna.UI2.WinForms;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ChiaseNoiBo
{
   
    public partial class Home1 : Form
    {
        private Guna.UI2.WinForms.Guna2MessageDialog guna2MessageDialog12;
        private string excelusername;
        private readonly UserControl_LoadFile userControl_LoadFile;
        private readonly HuongDanSuDungControl huongDanSuDungControl;
        private string userRole;
        private Guna.UI2.WinForms.Guna2Button currentActiveButton;
        public Home1()
        {
            InitializeComponent();
            guna2MessageDialog12 = new Guna.UI2.WinForms.Guna2MessageDialog();
            guna2MessageDialog12.Parent = this;

            userControl_LoadFile = new UserControl_LoadFile();
            //Nếu như là không phải Admin thì ẩn đi nút 
            guna2Button6.Visible = false;
            // Khởi tạo label hiển thị trạng thái


        }
        public Home1(string excelusername, string role)
        {
            InitializeComponent();
            this.excelusername = excelusername;
            this.userRole = role;
            guna2MessageDialog12 = new Guna.UI2.WinForms.Guna2MessageDialog();
            guna2MessageDialog12.Parent = this;
            userControl_LoadFile = new UserControl_LoadFile();
            
            // Phân quyền: Nếu là Admin thì hiện nút
            guna2Button6.Visible = role == "Admin";

            
        }
        public void LoadUserControl(UserControl uc)
        {
            panel2.Controls.Clear();
            uc.Dock = DockStyle.Fill;
            panel2.Controls.Add(uc);
        }
       
        private void ActivateButton(Guna.UI2.WinForms.Guna2Button clickedButton)
        {
            if (currentActiveButton != null)
            {
                // Reset lại nút trước đó
                currentActiveButton.FillColor = Color.Transparent;
                currentActiveButton.ForeColor = Color.DeepSkyBlue;
            }

            // Cập nhật nút hiện tại
            currentActiveButton = clickedButton;
            currentActiveButton.FillColor = Color.DeepSkyBlue;
            currentActiveButton.ForeColor = Color.Black;
        }

        private void Home1_Load(object sender, EventArgs e)
        {
        ;
            string role=
            label2.Text = excelusername+$"({userRole})";
            LoadUserControl(new HuongDanSuDungControl());
            //guna2Button3.Visible = false;
            //guna2ComboBox1.Visible = false;
            //guna2PictureBox1.Visible = false;

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            Login login = new Login();
            login.ShowDialog();
        }
        //Nut Excel
        private async void guna2Button3_Click(object sender, EventArgs e)
        {
            ActivateButton((Guna.UI2.WinForms.Guna2Button)sender);
            var uc = new UserControl_LoadFile();
            LoadUserControl(uc);
           await uc.LoadExcelFilesAsync();

        }
        //Nut van ban
        private async void guna2Button4_Click(object sender, EventArgs e)
        {
            ActivateButton((Guna.UI2.WinForms.Guna2Button)sender);
            var uc = new UserControl_LoadFile();
            LoadUserControl(uc);
            await uc.LoadWordAndPdfFilesAsync();


        }
        // tao bieu mau
        private void guna2Button6_Click(object sender, EventArgs e)
        {

        }

        private void guna2ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
           
            panel2.Controls.Clear();

            if (guna2ComboBox1.SelectedItem.ToString() == "Đơn xin nghỉ phép")
            {
                var control = new DonXinNghiPhepControl();
                control.Dock = DockStyle.Fill;
                panel2.Controls.Add(control);
            }
            else if (guna2ComboBox1.SelectedItem.ToString() == "Đề xuất tăng lương")
            {
                var control = new DeXuatTangLuongControl();
                control.Dock = DockStyle.Fill;
                panel2.Controls.Add(control);
            }
        }

        public void ShowMessage(string message, string title, Guna.UI2.WinForms.MessageDialogIcon icon)
        {
            guna2MessageDialog12.Icon = icon;
            guna2MessageDialog12.Show(message, title);
        }


  
        // Nút cập nhật phiên bản
        private async void guna2Button6_Click_1(object sender, EventArgs e)
        {
            try
            {
                var credential = GoogleCredential.FromFile(GoogleDriveUpdater.CredentialPath)
                    .CreateScoped(GoogleDriveUpdater.Scopes);
                var service = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = GoogleDriveUpdater.ApplicationName,
                });

                // Kiểm tra phiên bản
                string localVersion = btn_version.Text.Replace("Phiên bản ", "").Trim();
                string driveVersion = await userControl_LoadFile.CheckOnlineVersion(service);

                if (string.IsNullOrEmpty(driveVersion))
                {
                    guna2MessageDialog1.Show("Không thể lấy phiên bản từ Google Drive!", "Lỗi");
                    return;
                }

                if (localVersion == driveVersion)
                {
                    guna2MessageDialog1.Show($"Bạn đang sử dụng phiên bản mới nhất: {localVersion}", "Thông báo");
                    return;
                }

                // Thông báo có bản cập nhật mới
                ShowMessage($"Hệ thống đã có phiên bản mới! Vui lòng cập nhật để trải nghiệm chức năng mới nhất!", "Thông báo",MessageDialogIcon.Information);

                string downloadPath = null;
                bool updateSuccess = false;

                using (var updateForm = new UpdateNotificationForm())
                {
                    updateForm.StartPosition = FormStartPosition.CenterParent;

                    updateForm.WorkToDoAsync = async () =>
                    {
                        try
                        {
                            // Bước 1: Lấy tên file MSI
                            string latestMsiFileName = await userControl_LoadFile.GetMsiFileName(service, GoogleDriveUpdater.FolderId);
                            if (string.IsNullOrEmpty(latestMsiFileName))
                                throw new Exception("Không tìm thấy file MSI trên Google Drive!");

                            // Bước 2: Tải file về
                            downloadPath = Path.Combine(userControl_LoadFile.SharedDirectory, latestMsiFileName);
                            updateSuccess = await DownloadFileFromDrive(service, GoogleDriveUpdater.FolderId, latestMsiFileName, downloadPath);

                            if (!updateSuccess)
                                throw new Exception("Tải file thất bại!");
                        }
                        catch (Exception ex)
                        {
                            // Ghi lại lỗi để xử lý sau
                            updateForm.TaskException = ex;
                            updateSuccess = false;
                        }
                    };

                    updateForm.ShowDialog(this);

                    // Xử lý kết quả sau khi form đóng
                    if (updateForm.TaskException != null)
                    {
                        throw updateForm.TaskException;
                    }
                }

                // Nếu cập nhật thành công thì tiến hành cài đặt
                if (updateSuccess && File.Exists(downloadPath))
                {
                   // guna2MessageDialog1.Show("Cập nhật hoàn tất! Ứng dụng sẽ tự động cài đặt phiên bản mới.", "Thành công");
                    InstallMsi(downloadPath);
                }
            }
            catch (Exception ex)
            {
                guna2MessageDialog1.Show($"Lỗi cập nhật: {ex.Message}", "Lỗi");
            }
        }
        public void InstallMsi(string msiFilePath)
        {
            try
            {
                //guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Information;
                //guna2MessageDialog1.Show($"Bắt đầu cài đặt: {Path.GetFileName(msiFilePath)}", "Cài đặt");

                using (Process process = new Process())
                {
                    process.StartInfo.FileName = "msiexec";
                    process.StartInfo.Arguments = $"/i \"{msiFilePath}\" /qn /norestart";
                    process.StartInfo.UseShellExecute = false;


                    process.Start();
                    process.WaitForExit();

                    if (process.ExitCode == 0)
                    {
                        guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Information;
                        guna2MessageDialog1.Show("Cài đặt hoàn tất! Ứng dụng sẽ khởi động lại.", "Thành công");

                        string filedelete = msiFilePath.Replace("\\", @"\");


                        try
                        {
                            if (File.Exists(filedelete))
                            {
                                File.Delete(filedelete); // Không cần dùng Path.Combine vì filedelete đã là đường dẫn hoàn chỉnh
                                //ShowMessage("Xóa file thành công!", "Thông báo", MessageDialogIcon.Question);
                            }
                            else
                            {
                                //ShowMessage("Không tìm thấy file để xóa!", "Thông báo", MessageDialogIcon.Warning);
                            }
                        }
                        catch (Exception ex)
                        {
                            //ShowMessage($"Xóa file thất bại!\nChi tiết lỗi: {ex.Message}", "Lỗi", MessageDialogIcon.Error);
                        }


                        // Khởi động lại ứng dụng
                        RestartApplication();
                    }
                    else
                    {
                        guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                        guna2MessageDialog1.Show($"Cài đặt thất bại, mã lỗi: {process.ExitCode}", "Lỗi");
                    }
                }
            }
            catch (Exception ex)
            {
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                guna2MessageDialog1.Show($"Lỗi khi cài đặt MSI:\n{ex.Message}", "Lỗi");
            }
        }
       
        public void RestartApplication()
        {
            try
            {
                var exePath = Application.ExecutablePath;
                if (!File.Exists(exePath))
                {
                    guna2MessageDialog1.Show($"Không tìm thấy file exe tại: {exePath}", "Lỗi khởi động lại");
                    return;
                }

                Process.Start(new ProcessStartInfo
                {
                    FileName = exePath,
                    UseShellExecute = true
                });

                Application.Exit();
            }
            catch (Exception ex)
            {
                guna2MessageDialog1.Show($"Lỗi khi khởi động lại hệ thống:\n{ex.Message}", "Lỗi");
            }
        }

        public async Task<bool> DownloadFileFromDrive(DriveService service, string folderId, string fileName, string downloadPath)
        {
            try
            {
                var listRequest = service.Files.List();
                listRequest.Q = $"'{GoogleDriveHelper. FolderId}' in parents and name = \"{fileName}\"";
                listRequest.Fields = "files(id, name)";

                var files = await listRequest.ExecuteAsync();


                if (files.Files.Count == 0)
                {
                    guna2MessageDialog1.Parent = this;
                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                    guna2MessageDialog1.Show($"Không tìm thấy file {fileName} trên Google Drive!", "Lỗi");
                    return false;
                }

                var file = files.Files[0];
                var request = service.Files.Get(file.Id);

                using (var fileStream = new FileStream(downloadPath, FileMode.Create, FileAccess.Write))
                {
                    await request.DownloadAsync(fileStream);
                }
              
                    //guna2MessageDialog1.Parent = this;
                    //guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Information;
                    //guna2MessageDialog1.Show($"File {fileName} đã được tải về thành công!", "Thành công");
                return true;



            }
            catch (Exception ex)
            {
                guna2MessageDialog1.Parent = this;
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                guna2MessageDialog1.Show("Lỗi khi tải file từ Google Drive: " + ex.Message, "Lỗi");
                return false;
            }
        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            ActivateButton((Guna.UI2.WinForms.Guna2Button)sender);
            LoadUserControl(new HuongDanSuDungControl());
        }

        private void guna2Button6_Click_2(object sender, EventArgs e)
        {
            ActivateButton((Guna.UI2.WinForms.Guna2Button)sender);
            LoadUserControl(new DuyetaikhoanControl());
        }
    }
    }


using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using File = System.IO.File;
using Label = System.Windows.Forms.Label;

namespace ChiaseNoiBo
{
    public partial class UserControl_LoadFile : UserControl
    {
        public  string SharedDirectory = AppDomain.CurrentDomain.BaseDirectory;
        private GoogleDriveHelper gg=new GoogleDriveHelper();
      
        private static string FolderId = "15viUYINHRFLMIuCNVI4khVOHZgMf13jN"; // Thư mục chứa file Excel
       
       
        public UserControl_LoadFile()
        {

            InitializeComponent();
            flowLayoutPanel1.SizeChanged += flowLayoutPanel1_SizeChanged;
           

        }

        private  void UserControl_LoadFile_Load(object sender, EventArgs e)
        {
            if (this.ParentForm is Home1 homeForm)
            {
                guna2MessageDialog1.Parent = homeForm;
            }
          
        }
        private void flowLayoutPanel1_SizeChanged(object sender, EventArgs e)
        {
            foreach (Control ctrl in flowLayoutPanel1.Controls)
            {
                if (ctrl is Button btn)
                {
                    btn.Width = flowLayoutPanel1.ClientSize.Width - 25;
                }
            }
        }

        /// <summary>
        /// Lấy danh sách file Excel trong thư mục Google Drive và hiển thị lên panel1
        /// </summary>
        public async Task LoadExcelFilesAsync()
        {
            try
            {
                var service = GoogleDriveHelper.GetDriveService();
                var request = service.Files.List();
                request.Q = $"'{FolderId}' in parents and (mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType='application/vnd.ms-excel' or mimeType='text/csv')";

                request.Fields = "files(id, name)";

                var result = await request.ExecuteAsync();

                if (result.Files == null || result.Files.Count == 0)
                {
                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                    guna2MessageDialog1.Show("Không tìm thấy file Excel nào trong thư mục!", "Thông báo");
                    return;
                }

                flowLayoutPanel1.Controls.Clear(); // Xóa danh sách cũ


                // Sử dụng FlowLayoutPanel để tự động sắp xếp
                foreach (var file in result.Files.Where(f => f.Name != "DanhSachTaiKhoan.xlsx"))
                {
                    // Panel chứa cả tên file và icon
                    var filePanel = new Panel
                    {
                        Width = flowLayoutPanel1.ClientSize.Width - 25,
                        Height = 50,
                        BackColor = Color.WhiteSmoke,
                        Margin = new Padding(5),
                        Padding = new Padding(5),
                        Tag = file.Id
                    };

                    // Nút hiển thị tên file
                    var fileLabel = new Label
                    {
                        Text = file.Name,
                        AutoSize = false,
                        Width = filePanel.Width - 60, // Trừ phần icon
                        Height = 40,
                        Font = new Font("Segoe UI", 12),
                        TextAlign = ContentAlignment.MiddleLeft,
                        Dock = DockStyle.Left,
                        Cursor = Cursors.Hand,
                        Padding = new Padding(5, 10, 5, 5),
                        Tag = file.Id
                    };
                    fileLabel.Click += FileButton_Click;

                    // Icon tải xuống
                    var downloadIcon = new PictureBox
                    {
                        Image = Properties.Resources.icon_download, // Đảm bảo bạn thêm icon "download" vào Resources
                        SizeMode = PictureBoxSizeMode.Zoom,
                        Width = 32,
                        Height = 32,
                        Cursor = Cursors.Hand,
                        Dock = DockStyle.Right,
                        Margin = new Padding(5),
                        Tag = file // Truyền cả object để lấy ID & Name
                    };
                    downloadIcon.Click += async (s, e) =>
                    {
                        var picBox = s as PictureBox;
                        var driveFile = picBox.Tag as Google.Apis.Drive.v3.Data.File;

                        using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                        {
                            // Lấy tên gốc và đuôi file
                            string originalName = Path.GetFileNameWithoutExtension(driveFile.Name);
                            string extension = Path.GetExtension(driveFile.Name);

                            // Gợi ý lưu tại thư mục Documents
                            string folder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                            // Tạo tên gợi ý: Tên_1, Tên_2, ...
                            int index = 1;
                            string suggestedName;
                            do
                            {
                                suggestedName = $"{originalName}_{index}{extension}";
                                index++;
                            } while (File.Exists(Path.Combine(folder, suggestedName)));

                            saveFileDialog.InitialDirectory = folder;
                            saveFileDialog.FileName = suggestedName;
                            saveFileDialog.Filter = "Excel Files|*.xlsx;*.xls|CSV Files|*.csv|All Files|*.*";

                            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                try
                                {
                                    var requestDownload = service.Files.Get(driveFile.Id);
                                    var stream = new MemoryStream();
                                    await requestDownload.DownloadAsync(stream);

                                    // Sau khi người dùng chọn tên, vẫn kiểm tra trùng để thêm hậu tố nếu cần
                                    string selectedPath = Path.GetDirectoryName(saveFileDialog.FileName);
                                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(saveFileDialog.FileName);
                                    string fileExt = Path.GetExtension(saveFileDialog.FileName);

                                    string finalPath = Path.Combine(selectedPath, fileNameWithoutExt + fileExt);
                                    int conflictIndex = 1;

                                    while (File.Exists(finalPath))
                                    {
                                        finalPath = Path.Combine(selectedPath, $"{fileNameWithoutExt}_{conflictIndex}{fileExt}");
                                        conflictIndex++;
                                    }

                                    File.WriteAllBytes(finalPath, stream.ToArray());
                                   
                                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Information;
                                    guna2MessageDialog1.Show("Tải file thành công!", "Thông báo");
                                }
                                catch (Exception ex)
                                {
                                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                                    guna2MessageDialog1.Show(ex.Message, "Lỗi");
                                }
                            }
                        }
                    };


                    // Thêm vào panel
                    filePanel.Controls.Add(fileLabel);
                    filePanel.Controls.Add(downloadIcon);
                    flowLayoutPanel1.Controls.Add(filePanel);
                }
                flowLayoutPanel1.AutoScroll = true;
                
            }
            catch (Exception ex)
            {
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                guna2MessageDialog1.Show($"Lỗi khi tải danh sách file: {ex.Message}", "Lỗi");
            }

        }
          

    public async Task LoadWordAndPdfFilesAsync()
        {
            try
            {
                var service = GoogleDriveHelper.GetDriveService();
                var request = service.Files.List();

                // MIME types cho Word và PDF
                request.Q = $"'{FolderId}' in parents and (" +
                            "mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document' or " + // .docx
                            "mimeType='application/msword' or " + // .doc
                            "mimeType='application/pdf')"; // .pdf

                request.Fields = "files(id, name)";

                var result = await request.ExecuteAsync();

                if (result.Files == null || result.Files.Count == 0)
                {
                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                    guna2MessageDialog1.Show("Không tìm thấy file Word hoặc PDF nào trong thư mục!", "Thông báo");
                    return;
                }

                flowLayoutPanel1.Controls.Clear(); // Xóa danh sách cũ
                foreach (var file in result.Files)
                {
                    // Panel chứa mỗi file
                    var filePanel = new Panel
                    {
                        Width = flowLayoutPanel1.ClientSize.Width - 25,
                        Height = 50,
                        BackColor = Color.WhiteSmoke,
                        Margin = new Padding(5),
                        Padding = new Padding(5),
                        Tag = file.Id
                    };

                    // Tên file (Label clickable)
                    var fileLabel = new Label
                    {
                        Text = file.Name,
                        AutoSize = false,
                        Width = filePanel.Width - 60, // Trừ khoảng cho icon
                        Height = 40,
                        Font = new Font("Segoe UI", 12),
                        TextAlign = ContentAlignment.MiddleLeft,
                        Dock = DockStyle.Left,
                        Cursor = Cursors.Hand,
                        Padding = new Padding(5, 10, 5, 5),
                        Tag = file.Id
                    };
                    fileLabel.Click += FileButton_Click;

                    // Icon tải về (PictureBox)
                    var downloadIcon = new PictureBox
                    {
                        Image = Properties.Resources.icon_download, // Nhớ add icon này trong Resources
                        SizeMode = PictureBoxSizeMode.Zoom,
                        Width = 32,
                        Height = 32,
                        Cursor = Cursors.Hand,
                        Dock = DockStyle.Right,
                        Margin = new Padding(5),
                        Tag = file // Truyền full object để lấy ID & Name
                    };
                    downloadIcon.Click += async (s, e) =>
                    {
                        var picBox = s as PictureBox;
                        var driveFile = picBox.Tag as Google.Apis.Drive.v3.Data.File;

                        using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                        {
                            // Tách tên file và extension
                            string originalName = Path.GetFileNameWithoutExtension(driveFile.Name);
                            string extension = Path.GetExtension(driveFile.Name);

                            // Đặt tên gốc là tên_1
                            string baseName = originalName;
                            string suggestedName = baseName + extension;
                            string folder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); // Hoặc mặc định theo nhu cầu

                            int index = 1;
                            while (File.Exists(Path.Combine(folder, suggestedName)))
                            {
                                index++;
                                suggestedName = $"{originalName}_{index}{extension}";
                            }

                            // Đề xuất file chưa tồn tại
                            saveFileDialog.InitialDirectory = folder;
                            saveFileDialog.FileName = suggestedName;
                            saveFileDialog.Filter = "Excel Files|*.xlsx;*.xls|CSV Files|*.csv|All Files|*.*";

                            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                try
                                {
                                    var requestDownload = service.Files.Get(driveFile.Id);
                                    var stream = new MemoryStream();
                                    await requestDownload.DownloadAsync(stream);

                                    File.WriteAllBytes(saveFileDialog.FileName, stream.ToArray());

                                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Information;
                                    guna2MessageDialog1.Show("Tải file thành công!", "Thông báo");
                                }
                                catch (Exception ex)
                                {
                                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                                    guna2MessageDialog1.Show(ex.Message, "Lỗi");
                                }
                            }
                        }
                    };


                    // Gắn vào FlowLayoutPanel
                    filePanel.Controls.Add(fileLabel);
                    filePanel.Controls.Add(downloadIcon);
                    flowLayoutPanel1.Controls.Add(filePanel);
                }

                flowLayoutPanel1.AutoScroll = true;
            }
            catch (Exception ex)
            {
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                guna2MessageDialog1.Show($"Lỗi khi tải danh sách file Word/PDF: {ex.Message}", "Lỗi");
            }
        }

        private void FileButton_Click(object sender, EventArgs e)
        {
            Control control = sender as Control;
            if (control == null || control.Tag == null) return;
            string fileId = control.Tag.ToString();
            string fileName = control.Text;


            // Lấy phần mở rộng file (không phân biệt hoa thường)
            string extension = Path.GetExtension(fileName).ToLower();

            if (extension == ".xlsx" || extension == ".xls" || extension == ".csv")
            {
                // Nếu là file Excel hoặc CSV => mở Form3
                Form3_LoadExcel form = new Form3_LoadExcel(fileId, fileName);
                form.Show();
            }
            else
            {
                // Mặc định mở Form1
                Form1 form = new Form1(fileId, fileName);
                form.Show();
            }
        }


        public async Task<string> CheckOnlineVersion(DriveService service)
        {
            try
            {
                var fileMetadata = await service.Files.Get(GoogleDriveUpdater.VersionFileId).ExecuteAsync();

                if (fileMetadata.MimeType == "application/vnd.google-apps.document")
                {
                    var request = service.Files.Export(GoogleDriveUpdater.VersionFileId, "text/plain");
                    using (var stream = new MemoryStream())
                    {
                        await request.DownloadAsync(stream);
                        stream.Position = 0;
                        using (var reader = new StreamReader(stream))
                        {
                            string firstLine = reader.ReadLine();
                            return firstLine?.Trim() ?? "Error";
                        }
                    }
                }
                else if (fileMetadata.MimeType == "text/plain")
                {
                    var request = service.Files.Get(GoogleDriveUpdater.VersionFileId);
                    using (var stream = new MemoryStream())
                    {
                        await request.DownloadAsync(stream);
                        stream.Position = 0;
                        using (var reader = new StreamReader(stream))
                        {
                            string firstLine = reader.ReadLine();
                            return firstLine?.Trim() ?? "Error";
                        }
                    }
                }
                else
                {
                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                    guna2MessageDialog1.Show("File không đúng định dạng TXT hoặc Google Docs!", "Lỗi");
                    return "Error";
                }
            }
            catch (Exception ex)
            {
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                guna2MessageDialog1.Show("Lỗi khi tải file: " + ex.Message, "Lỗi");
                return "Error";
            }
        }

        public async Task<bool> DownloadFileFromDrive(DriveService service, string folderId, string fileName, string downloadPath)
        {
            try
            {
                var listRequest = service.Files.List();
                listRequest.Q = $"'{FolderId}' in parents and name = \"{fileName}\"";
                listRequest.Fields = "files(id, name)";

                var files = await listRequest.ExecuteAsync();
             

                if (files.Files.Count == 0)
                {
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
                Form parentForm = this.FindForm();
                if (parentForm != null)
                {
                    guna2MessageDialog1.Parent = parentForm;
                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Information;
                    guna2MessageDialog1.Show($"File {fileName} đã được tải về thành công!", "Thành công");
                }
                else
                {
                    MessageBox.Show($"File {fileName} đã được tải về thành công!", "Thành công");
                }

                return true;
               
            }
            catch (Exception ex)
            {
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                guna2MessageDialog1.Show("Lỗi khi tải file từ Google Drive: " + ex.Message, "Lỗi");
                return false;
            }
        }


        public async Task<string> GetMsiFileName(DriveService service, string folderId)
        {
            try
            {
                
                var listRequest = service.Files.List();
                listRequest.Q = $"'{folderId}' in parents and name contains '.msi'";  // Tìm các file .msi
                listRequest.Fields = "files(id, name)";  // Lấy thông tin id và name của file

                var files = await listRequest.ExecuteAsync();

                if (files.Files.Count == 0)
                {
                    return null; // Không tìm thấy file MSI nào
                }

                // Nếu có file MSI, trả về tên của file đầu tiên tìm thấy
                return files.Files[0].Name;
            }
            catch (Exception ex)
            {
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                guna2MessageDialog1.Show("Lỗi khi lấy file MSI: " + ex.Message, "Lỗi");
                return null;
            }
        }

        public void InstallMsi(string msiFilePath)
        {
            try
            {
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Information;
                guna2MessageDialog1.Show($"Bắt đầu cài đặt: {Path.GetFileName(msiFilePath)}", "Cài đặt");

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


                        // Xóa file MSI sau khi cài đặt thành công
                        if (File.Exists(filedelete))
                        {
                            File.Delete(Path.Combine(filedelete));
                        
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
            Process.Start(new ProcessStartInfo
            {
                FileName = Application.ExecutablePath,
                UseShellExecute = true
            });

            Application.Exit();
        }

        private async void guna2Button2_Click(object sender, EventArgs e)
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

                string localVersion = guna2Button2.Text.Replace("Phiên bản ", "").Trim();
                string driveVersion = await CheckOnlineVersion(service);

                if (string.IsNullOrEmpty(driveVersion))
                {
                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                    guna2MessageDialog1.Show("Không thể lấy phiên bản từ Google Drive!", "Lỗi");
                   
                    return;
                }

                if (localVersion == driveVersion)
                {
                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Information;
                    guna2MessageDialog1.Show($"Bạn đang sử dụng phiên bản mới nhất: {localVersion}", "Thông báo");
                  
                    return;
                }

                // Thông báo có phiên bản mới
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Warning;
                guna2MessageDialog1.Show($"Hệ thống đã có phiên bản mới! Vui lòng cập nhật để trải nghiệm chức năng mới nhất!");

                // Bắt đầu tải file
                string latestMsiFileName = await GetMsiFileName(service, GoogleDriveUpdater.FolderId);
                if (string.IsNullOrEmpty(latestMsiFileName))
                {
                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                    guna2MessageDialog1.Show("Không tìm thấy file MSI trên Google Drive!", "Lỗi");
                    
                    return;
                }

                string downloadPath = Path.Combine(SharedDirectory, latestMsiFileName);
                bool downloadSuccess = await DownloadFileFromDrive(service, GoogleDriveUpdater.FolderId, latestMsiFileName, downloadPath);
                if (!downloadSuccess)
                {
                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                    guna2MessageDialog1.Show("Không thể tải file cập nhật từ Google Drive!", "Lỗi");
                    
                    return;
                }

                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Information;
                guna2MessageDialog1.Show("Cập nhật hoàn tất! Ứng dụng sẽ tự động cài đặt phiên bản mới.", "Cập nhật thành công");

                // Cài đặt file MSI
                InstallMsi(downloadPath);

            }
            catch (Exception ex)
            {
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                guna2MessageDialog1.Show($"Lỗi khi kiểm tra/cập nhật phiên bản: {ex.Message}", "Lỗi");
               
            }
        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}

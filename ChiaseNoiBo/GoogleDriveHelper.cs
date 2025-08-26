using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Drive.v3.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MimeKit;
using MailKit.Net.Smtp;
using OfficeOpenXml;

namespace ChiaseNoiBo
{
    internal class GoogleDriveHelper
    {
        private const string CredentialPath = "credentials.json"; // Đường dẫn credentials.json
        private const string ApplicationName = "GoogleDriveUploader";
        public static string FolderId { get; } = "15viUYINHRFLMIuCNVI4khVOHZgMf13jN"; // Thư mục Google Drive chứa file
        private static readonly string[] Scopes = { DriveService.Scope.Drive };
        private static List<string> previousFileIds = new List<string>(); // Danh sách file đã kiểm tra trước đó

        public static DriveService GetDriveService()
        {
            try
            {
                var credential = GoogleCredential.FromFile(CredentialPath).CreateScoped(Scopes);
                return new DriveService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Lỗi khi khởi tạo DriveService: {ex.Message}");
                return null;
            }
        }

        public static List<Google.Apis.Drive.v3.Data.File> GetSpreadsheetFiles()
        {
            var service = GetDriveService();
            if (service == null) return new List<Google.Apis.Drive.v3.Data.File>();

            var request = service.Files.List();
            request.Q = $"'{FolderId}' in parents and (mimeType='application/vnd.ms-excel' " +
                        "or mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' " +
                        "or mimeType='text/csv')";
            request.Fields = "files(id, name)";

            try
            {
                var result = request.Execute();
                return result.Files as List<Google.Apis.Drive.v3.Data.File>;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Lỗi khi lấy danh sách file: {ex.Message}");
                return new List<Google.Apis.Drive.v3.Data.File>();
            }
        }
        //Gui mail khi co file moi
    //    public static void CheckForNewFiles()
    //    {
    //        var files = GetSpreadsheetFiles();
    //        var newFiles = files.Where(f => !previousFileIds.Contains(f.Id)).ToList();

    //        if (newFiles.Any())
    //        {
    //            string message = "📢 Hệ thống thông báo:\n";
    //            foreach (var file in newFiles)
    //            {
    //                message += $"- File mới: {file.Name}\n";
    //            }

    //            SendEmailNotification(message);
    //            previousFileIds = files.Select(f => f.Id).ToList();
    //        }
    //    }

    //    public static void SendEmailNotification(string message)
    //    {
    //        string smtpServer = "smtp.gmail.com";
    //        int port = 587;
    //        string fromEmail = "huynhanhkhoa30042019@gmail.com";
    //        string password = Environment.GetEnvironmentVariable("GMAIL_APP_PASSWORD"); // Dùng biến môi trường thay vì hardcode
    //        if (string.IsNullOrEmpty(password))
    //        {
    //            Console.WriteLine("❌ Lỗi: Mật khẩu Gmail không được cấu hình trong biến môi trường.");
    //            return;
    //        }

    //        string excelFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DanhSachTaiKhoan.xlsx");
    //        List<string> emailList = ExcelHelper.GetEmailsFromExcel(excelFilePath);

    //        if (!emailList.Any())
    //        {
    //            Console.WriteLine("❌ Không có email nào trong danh sách.");
    //            return;
    //        }

    //        try
    //        {
    //            var emailMessage = new MimeMessage();
    //            emailMessage.From.Add(new MailboxAddress("Hệ thống", fromEmail));
    //            emailList.ForEach(email => emailMessage.To.Add(new MailboxAddress(email, email)));

    //            emailMessage.Subject = "📢 Thông báo: File mới trong thư mục Google Drive";
    //            emailMessage.Body = new TextPart("plain")
    //            {
    //                Text = $"Xin chào,\n\n"
    //                     + $"Hệ thống có một thông báo mới 📢. Hãy vào xem để cập nhật thông tin mới nhất!\n\n"
    //                     + $"📂 Chi tiết:\n"
    //                     + $"{message}\n\n"
    //                     + "Vui lòng kiểm tra ngay!\n\n"
    //                     + "Trân trọng,\n"
    //                     + "Hệ thống thông báo"
    //            };

    //            using (var client = new SmtpClient())
    //            {
    //                client.Connect(smtpServer, port, false);
    //                client.Authenticate(fromEmail, password);
    //                client.Send(emailMessage);
    //                client.Disconnect(true);
    //            }

    //            Console.WriteLine("✅ Email thông báo đã được gửi!");
    //        }
    //        catch (Exception ex)
    //        {
    //            Console.WriteLine($"❌ Lỗi khi gửi email: {ex.Message}");
    //        }
    //    }
    //}

    //internal class ExcelHelper
    //{
    //    public static List<string> GetEmailsFromExcel(string filePath)
    //    {
    //        var emailList = new List<string>();

    //        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Để tránh lỗi bản quyền

    //        if (!System.IO.File.Exists(filePath))
    //        {
    //            Console.WriteLine("❌ Không tìm thấy file Excel!");
    //            return emailList;
    //        }

    //        try
    //        {
    //            using (var package = new ExcelPackage(new FileInfo(filePath)))
    //            {
    //                var worksheet = package.Workbook.Worksheets[0]; // Lấy sheet đầu tiên
    //                int rowCount = worksheet.Dimension.Rows;

    //                for (int row = 2; row <= rowCount; row++) // Bỏ qua tiêu đề (bắt đầu từ dòng 2)
    //                {
    //                    string email = worksheet.Cells[row, 3].Text.Trim();
    //                    if (!string.IsNullOrEmpty(email))
    //                    {
    //                        emailList.Add(email);
    //                    }
    //                }
    //            }
    //        }
    //        catch (Exception ex)
    //        {
    //            Console.WriteLine($"❌ Lỗi khi đọc file Excel: {ex.Message}");
    //        }

    //        return emailList;
    //    }
    }
}


using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;

namespace ChiaseNoiBo
{
    class GoogleDriveUpdater
    {
        public static string[] Scopes = { DriveService.Scope.Drive };
        public const string ApplicationName = "Schedule a medical examination";
        public const string FolderId = "15viUYINHRFLMIuCNVI4khVOHZgMf13jN"; // ID thư mục chứa file .msi
        public const string VersionFileId = "1X9yfvSSMx8KsjPDDWwVupHtlowHUeifE"; // ID của file version.txt
        public static string CredentialPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "credentials.json");

        
    }
}


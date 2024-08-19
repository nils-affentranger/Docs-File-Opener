using System;
using System.IO;
using System.Threading;
using System.Diagnostics;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;

class Program
{
    static string[] Scopes = { DriveService.Scope.Drive };
    static string ApplicationName = "Google Docs Uploader";

    static void Main(string[] args)
    {
        UserCredential credential;

        using (var stream = new FileStream(@"C:\Users\Nils A. (GIBZ)\OneDrive - GIBZ\Dokumente\client_secret_197223837930-6e95ar0jkaasceo1rkdvr4ovnoqjdkf7.apps.googleusercontent.com.json", FileMode.Open, FileAccess.Read))
        {
            string credPath = "token.json";
            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.FromStream(stream).Secrets,
                Scopes,
                "user",
                CancellationToken.None,
                new FileDataStore(credPath, true)).Result;
            Console.WriteLine("Credential file saved to: " + credPath);
        }

        // Create Drive API service.
        var service = new DriveService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = ApplicationName,
        });

        string folderPath = @"C:\Users\Nils A. (GIBZ)\OneDrive - GIBZ\Dokumente\Import"; // Replace with your folder path
        var watcher = new FileSystemWatcher(folderPath);
        watcher.Created += (sender, e) => OnFileCreated(sender, e, service);
        watcher.EnableRaisingEvents = true;

        Console.WriteLine("Watching folder for new document files. Press 'q' to quit.");
        while (Console.Read() != 'q') ;
    }

    static void OnFileCreated(object sender, FileSystemEventArgs e, DriveService service)
    {
        if (IsDocumentFile(e.FullPath))
        {
            Console.WriteLine($"New document detected: {e.Name}");
            UploadToGoogleDrive(e.FullPath, service);
        }
    }

    static bool IsDocumentFile(string filePath)
    {
        string extension = Path.GetExtension(filePath).ToLower();
        return Array.IndexOf(new string[] {
            ".txt", ".rtf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
            ".odt", ".ods", ".odp", ".pdf"
        }, extension) >= 0;
    }

    static void UploadToGoogleDrive(string filePath, DriveService service)
    {
        try
        {
            var fileMetadata = new Google.Apis.Drive.v3.Data.File()
            {
                Name = Path.GetFileName(filePath),
                MimeType = GetGoogleDocsMimeType(filePath)
            };

            FilesResource.CreateMediaUpload request;
            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                request = service.Files.Create(fileMetadata, stream, GetMimeType(filePath));
                request.Fields = "id, webViewLink";
                request.Upload();
            }

            var file = request.ResponseBody;
            Console.WriteLine($"File ID: {file.Id}");
            Console.WriteLine($"Web View Link: {file.WebViewLink}");

            // Open the file in browser
            Process.Start(new ProcessStartInfo(file.WebViewLink) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    static string GetMimeType(string fileName)
    {
        string extension = Path.GetExtension(fileName).ToLower();
        switch (extension)
        {
            // Text documents
            case ".txt": return "text/plain";
            case ".rtf": return "application/rtf";
            case ".pdf": return "application/pdf";

            // Microsoft Office formats
            case ".doc": return "application/msword";
            case ".docx": return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            case ".xls": return "application/vnd.ms-excel";
            case ".xlsx": return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            case ".ppt": return "application/vnd.ms-powerpoint";
            case ".pptx": return "application/vnd.openxmlformats-officedocument.presentationml.presentation";

            // OpenDocument formats
            case ".odt": return "application/vnd.oasis.opendocument.text";
            case ".ods": return "application/vnd.oasis.opendocument.spreadsheet";
            case ".odp": return "application/vnd.oasis.opendocument.presentation";

            // Default
            default: return "application/octet-stream";
        }
    }

    static string GetGoogleDocsMimeType(string fileName)
    {
        string extension = Path.GetExtension(fileName).ToLower();
        switch (extension)
        {
            // Convert to Google Docs
            case ".txt":
            case ".rtf":
            case ".doc":
            case ".docx":
            case ".odt":
                return "application/vnd.google-apps.document";

            // Convert to Google Sheets
            case ".xls":
            case ".xlsx":
            case ".ods":
                return "application/vnd.google-apps.spreadsheet";

            // Convert to Google Slides
            case ".ppt":
            case ".pptx":
            case ".odp":
                return "application/vnd.google-apps.presentation";

            // Keep as PDF
            case ".pdf":
                return "application/pdf";

            // Default: Store as Google Drive file without conversion
            default:
                return "application/vnd.google-apps.file";
        }
    }
}
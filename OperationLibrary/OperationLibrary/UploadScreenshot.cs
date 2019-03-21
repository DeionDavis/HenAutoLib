using System;
using System.Windows.Forms;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;

namespace OperationLibrary
{
    class UploadScreenshot
    {
        static string[] Scopes = { DriveService.Scope.Drive };
        static string ApplicationName = "Upload_Screenshot_Drive";
        public string UploadImage(string path)
        {
            
            string fid;
            UserCredential credential;
            credential = GetCredentials();
            // Create Drive API service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
           var link= ScrenshotUpload(path, service);
            return link;
        }
        //uploadng failed screenshot
        private string  ScrenshotUpload(string path, DriveService service)
        {
            var fileMetadata = new Google.Apis.Drive.v3.Data.File();
            fileMetadata.Name = Path.GetFileName(path);
            fileMetadata.MimeType = "image/*";
            FilesResource.CreateMediaUpload request;
            using (var stream = new System.IO.FileStream(path, System.IO.FileMode.Open))
            {
                request = service.Files.Create(fileMetadata, stream, "image/*");
                request.Fields = "id";
                request.Upload();
            }
            string fid = request.ResponseBody.Id;
            //writting permissions to access image
            Permission newPermission = new Permission();
            newPermission.Type = "anyone";
            newPermission.Role = "reader";
            service.Permissions.Create(newPermission, fid).Execute();
            var link = "https://drive.google.com/file/d/" + fid + "/view?usp=sharing";
            return link;
        }
        //getting permissions
        private UserCredential GetCredentials()
        {
            UserCredential credential;
            using (var stream = new FileStream("D:\\HENKEL_AUTOMATION\\Henkel_Projects\\LibraryandExecution\\OperationLibrary\\OperationLibrary\\client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, "client_secreta.json");
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            return credential;
        }
    }
       
    }


using System;
using Google.Apis.Drive.v3;
using Google.Apis.Auth.OAuth2;
using System.Threading;
using Google.Apis.Util.Store;
using Google.Apis.Services;
using Google.Apis.Drive.v3.Data;
using System.Collections;
using System.Xml.Linq;

namespace SourceCode.SmartObjects.Services
{
    public class GoogleDriveHelper
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/drive-dotnet-quickstart.json
        static string[] Scopes = { DriveService.Scope.DriveFile };
        static string ApplicationName = "Drive API .NET Quickstart";

        private static UserCredential GetUserCredential(string pCredentialPath)
        {
            UserCredential credential;

            //using (var stream = new System.IO.FileStream(pCredentialPath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
            //{
            //    //string credPath = System.IO.Path.GetDirectoryName(pCredentialPath);
            //    string credPath = System.Environment.GetFolderPath(
            //        System.Environment.SpecialFolder.Personal);
            //    credPath = System.IO.Path.Combine(credPath, ".credentials/drive-dotnet-quickstart.json");

            //    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
            //        GoogleClientSecrets.Load(stream).Secrets,
            //        Scopes,
            //        "user",
            //        CancellationToken.None,
            //        new FileDataStore(credPath, true)).Result;
            //}
            using (var stream =
                new System.IO.FileStream("client_secret.json", System.IO.FileMode.Open, System.IO.FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);

                //Andy ZHANG
                credPath = System.IO.Path.Combine(credPath, ".credentials\\drive-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }
            return credential;

        }

        public static DriveService GetDriveService(string pCredentialPath)
        {
            UserCredential credential = GetUserCredential(pCredentialPath);

            // Create Drive API service.
            DriveService service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }

        //// 

        /// Create a new Directory.
        /// Documentation: https://developers.google.com/drive/v3/web/folder
        /// 
        /// a Valid authenticated DriveService
        /// The title of the file. Used to identify file or folder name.
        /// A short description of the file.
        /// Collection of parent folders which contain this file. 
        /// Setting this field will put the file in all of the provided folders. root folder.
        /// 
        public static File CreateDirectory(DriveService _service, string _name, string _description, string _parent)
        {

            File NewDirectory = null;
            try
            {
                // Create metaData for a new Directory
                var fileMetadata = new File()
                {
                    Name = _name,
                    Description = _description,
                    MimeType = "application/vnd.google-apps.folder",
                    Parents = new System.Collections.Generic.List<string>() { _parent }
                };
                var request = _service.Files.Create(fileMetadata);
                request.Fields = "id";
                NewDirectory = request.Execute();
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
            }

            return NewDirectory;
        }

        // tries to figure out the mime type of the file.
        private static string GetMimeType(string fileName)
        {
            string mimeType = "application/unknown";
            string ext = System.IO.Path.GetExtension(fileName).ToLower();
            Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(ext);
            if (regKey != null && regKey.GetValue("Content Type") != null)
                mimeType = regKey.GetValue("Content Type").ToString();
            return mimeType;
        }

        /// 
        /// Uploads a file
        /// Documentation: https://developers.google.com/drive/v3/web/manage-uploads
        /// 
        /// a Valid authenticated DriveService
        /// path to the file to upload
        /// Collection of parent folders which contain this file. 
        /// Setting this field will put the file in all of the provided folders. root folder.
        /// If upload succeeded returns the File resource of the uploaded file 
        /// If the upload fails returns null
        public static File UploadFile(DriveService _service, string _uploadFile, string _parent)
        {
            // Parse original file name and also document base64 string
            XDocument xDoc = new XDocument();
            xDoc = XDocument.Parse(_uploadFile);
            string strFileName = xDoc.Root.Element("name").Value;
            byte[] bFile = Convert.FromBase64String(xDoc.Root.Element("content").Value);
            string strFileExt = System.IO.Path.GetExtension(strFileName);
            string strMimeType = GetMimeType(strFileName);
            var fileMetadata = new File()
            {
                Name = strFileName,
                MimeType= strMimeType,
                Parents = new System.Collections.Generic.List<string>() { _parent }

            };
            File file = null;
            using (System.IO.Stream stream = new System.IO.MemoryStream(bFile))
            {
                FilesResource.CreateMediaUpload request;
                request = _service.Files.Create(fileMetadata, stream, strMimeType);
                request.Fields = "id";
                request.Upload();
                file = request.ResponseBody;
            }
                
            return file;
            
        }

        public static byte[] DownloadbtFile(DriveService _service, string _fileId)
        {
            var stream = new System.IO.MemoryStream();
            var request = _service.Files.Get(_fileId);
            request.Download(stream);
            byte[] btfile = stream.ToArray();
            //var str = Convert.ToBase64String(btfile);
            //memoryStream.CopyTo(fileStream);
            System.IO.File.WriteAllBytes(@"C:\Download\download.pdf", btfile);
            return btfile;
        }

        public static string DownloadFile(DriveService _service, string _fileId)
        {
            var stream = new System.IO.MemoryStream();
            var request = _service.Files.Get(_fileId);
            request.Download(stream);
            byte[] btfile = stream.ToArray();
            var str = Convert.ToBase64String(btfile);
            return str;
        }
    }
}

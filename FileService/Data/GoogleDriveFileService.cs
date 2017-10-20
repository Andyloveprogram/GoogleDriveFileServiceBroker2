using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Linq;
using SIO = System.IO;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using Attributes = SourceCode.SmartObjects.Services.ServiceSDK.Attributes;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Threading;
using Google.Apis.Download;
using Google.Apis.Drive.v2.Data;
using Google.Apis.Drive.v3.Data;

namespace SourceCode.SmartObjects.Services
{
    /// <summary>
    /// Sample implementation of a Static Service Object.
    /// The class is decorated with a ServiceObject Attribute providing definition information for the Service Object.
    /// The Properties and Methods in the class are each decorated with attributes that describe them in Service Object terms
    /// This sample implementation contains two Properties (Number and Text) and two Methods (Read and List)
    /// </summary>
    [Attributes.ServiceObject("GoogleDriveFileService", "Google Drive File Service", "Google Drive File Service")]
    class GoogleDriveFileService
    {
        

        /// <summary>
        /// This property is required if you want to get the service instance configuration 
        /// settings in this class
        /// </summary>
        private ServiceConfiguration _serviceConfig;
        public ServiceConfiguration ServiceConfiguration
        {
            get { return _serviceConfig; }
            set { _serviceConfig = value; }
        }

        #region Class Level Fields

        #region Private Fields
        private const string credName = "OAuth 2.0 Client ID Path";


        #endregion

        #endregion

        #region Properties with Property Attribute

        #region string FileData
        /// <summary>
        /// The File object representing the file being uploaded.
        /// The property is decorated with a Property Attribute which describes the Service Object Property
        /// </summary>
        [Attributes.Property("FileData", SoType.File, "File Data", "File Data")]
        public string FileData
        {
            get;
            set;
        }
        #endregion

        #region string FileName
        /// <summary>
        /// An string value representing the desired file name.
        /// The property is decorated with a Property Attribute which describes the Service Object Property
        /// </summary>
        [Attributes.Property("FileName", SoType.Text, "File Name", "File Name")]
        public string FileName
        {
            get;
            set;
        }

        [Attributes.Property("FileId", SoType.Text, "File Id", "File Id")]
        public string FileId
        {
            get;
            set;
        }
        #endregion

        #region string FilePath
        /// <summary>
        /// An string value representing the uploaded location file path.
        /// The property is decorated with a Property Attribute which describes the Service Object Property
        /// </summary>
        //[Attributes.Property("FilePath", SoType.Text, "File Path", "File Path")]
        //public string FilePath
        //{
        //    get { return pFilePath; }
        //    set { pFilePath = value; }
        //}
        #endregion

        #region string FolderPath
        /// <summary>
        /// An string value representing the uploaded location file path.
        /// The property is decorated with a Property Attribute which describes the Service Object Property
        /// </summary>
        [Attributes.Property("FolderPath", SoType.Text, "Folder Path", "Folder Path")]
        public string FolderPath
        {
            get;
            set;
        }

        [Attributes.Property("FolderName", SoType.Text, "Folder Name", "Folder Name")]
        public string FolderName
        {
            get;
            set;
        }
        #endregion

        #endregion


        #region Default Constructor
        /// <summary>
        /// Instantiates a new ServiceObject1.
        /// </summary>
        public GoogleDriveFileService()
        {
            // No implementation necessary.
        }
        #endregion

        #region Methods with Method Attribute

        #region FileService Create(Property[] inputs, RequiredProperties required, Property[] returns, MethodType methodType, ServiceObject serviceObject)
        /// <summary>
        /// Sample implementation of a static service Object Read method
        /// The method is decorated with a Method Attribute which describes the Service Object Method.
        /// </summary>
        //[Attributes.Method("Create", MethodType.Execute, "Create", "Creates a new file.",
        //    new string[] { "FileData", "FolderPath" }, //required property array. Properties must match the names of Properties decorated with the Property Attribute
        //    new string[] { "FileData", "FolderPath" }, //input property array. Properties must match the names of Properties decorated with the Property Attribute
        //    new string[] { "FileId" })] //return property array. Properties must match the names of Properties decorated with the Property Attribute
        //public GoogleDriveFileService Create()
        //{
        //    // Parse original file name and also document base64 string
        //    XDocument xDoc = new XDocument();
        //    xDoc = XDocument.Parse(this.FileData);
        //    string sFileName = xDoc.Root.Element("name").Value;
        //    byte[] bFile = Convert.FromBase64String(xDoc.Root.Element("content").Value);
        //    // return the full path
        //    return this;
        //}
        #endregion

        [Attributes.Method("List", MethodType.List, "List", "List files",
            new string[] { "FolderPath" }, //required property array. Properties must match the names of Properties decorated with the Property Attribute
            new string[] { "FolderPath" }, //input property array. Properties must match the names of Properties decorated with the Property Attribute
            new string[] { "FileName", "FolderPath" ,"FileId", "FileData" })] //return property array. Properties must match the names of Properties decorated with the Property Attribute
        public List<GoogleDriveFileService> List()
        {
            string credPath = this.ServiceConfiguration[credName].ToString();
            var service = GoogleDriveHelper.GetDriveService(credPath);

            List<GoogleDriveFileService> listFile = new List<GoogleDriveFileService>();

            // Define parameters of request.
            
            Google.Apis.Drive.v3.FilesResource.ListRequest listRequest = service.Files.List();
            listRequest.Q = string.Format("'{0}' in parents and trashed=false",this.FolderPath);
            //listRequest.PageSize = 10;
            listRequest.Fields = "files";            
            // List files.
            IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute().Files;
            Console.WriteLine("Files:");
            if (files != null && files.Count > 0)
            {
                foreach (var file in files)
                {
                    GoogleDriveFileService fService = new GoogleDriveFileService();
                    string strFile = GoogleDriveHelper.DownloadFile(service,file.Id);
                    StringBuilder sb = new StringBuilder();
                    sb.Append("<file>");
                    sb.AppendFormat("<name>{0}</name>", file.Name);
                    sb.AppendFormat("<content>{0}</content>", strFile);
                    sb.Append("</file>");

                    fService.FileData = sb.ToString();
                    fService.FileName = file.Name;
                    fService.FileId = file.Id;
                    if(file.Parents!=null && file.Parents.Count>=1)
                    {
                        fService.FolderPath = file.Parents[0].ToString();
                    }
                    
                    listFile.Add(fService);
                    Console.WriteLine("{0} ({1})", file.Name, file.Id);
                }
            }
            else
            {
                Console.WriteLine("No files found.");
            }
            return listFile;

        }
        //Andy ZHANG
        //@2017-7-13
        //Upload File Method to Google Drive 
        [Attributes.Method("Create", MethodType.Execute, "Upload File", "Upload File",
                    new string[] { "FileData", "FolderPath" }, //required property array. Properties must match the names of Properties decorated with the Property Attribute
                    new string[] { "FileData", "FolderPath" }, //input property array. Properties must match the names of Properties decorated with the Property Attribute
                    new string[] { "FileId" })] //return property array. Properties must match the names of Properties decorated with the Property Attribute
        public GoogleDriveFileService UploadFile()
        {
            // Parse original file name and also document base64 string
            XDocument xDoc = new XDocument();
            xDoc = XDocument.Parse(this.FileData);
            string strFileName = xDoc.Root.Element("name").Value;
            byte[] bFile = Convert.FromBase64String(xDoc.Root.Element("content").Value);
            string strFileExt = System.IO.Path.GetExtension(strFileName);
            string strMimeType = GetMimeType(strFileName);
            var fileMetadata = new Google.Apis.Drive.v3.Data.File()
            {
                Name = strFileName,
                MimeType = strMimeType,
                Parents = new List<string>
                    {
                        this.FolderPath
                    },
                

            };
            Google.Apis.Drive.v3.Data.File file = null;
            using (System.IO.Stream stream = new System.IO.MemoryStream(bFile))
            {
                Google.Apis.Drive.v3.FilesResource.CreateMediaUpload request;
                string credPath = this.ServiceConfiguration[credName].ToString();
                var _service = GoogleDriveHelper.GetDriveService(credPath);

                request = _service.Files.Create(fileMetadata, stream, strMimeType);
                request.Fields = "id";
                request.Upload();
                file = request.ResponseBody;
                if(file!=null&& !string.IsNullOrEmpty(file.Id))
                {
                    this.FileId = file.Id;
                }
            }
            //this.FileName= strFileName;
            return this;

        }

        

        //Andy ZHANG
        //@2017-7-13
        //Download File from google drive
        [Attributes.Method("Execute", MethodType.Execute, "DownloadFile", "DownloadFile",
                    new string[] { "FileId", "FileName" }, //required property array. Properties must match the names of Properties decorated with the Property Attribute
                    new string[] { "FileId","FileName" }, //input property array. Properties must match the names of Properties decorated with the Property Attribute
                    new string[] { "FileData" })] //return property array. Properties must match the names of Properties decorated with the Property Attribute
        public GoogleDriveFileService DownloadFile()
        {
            var fileId = this.FileId;
            string credPath = this.ServiceConfiguration[credName].ToString();
            var service = GoogleDriveHelper.GetDriveService(credPath);

            var stream = new System.IO.MemoryStream();
            var request = service.Files.Get(fileId);
            request.Download(stream);
            byte[] btfile = stream.ToArray();
            string strFile = Convert.ToBase64String(btfile);

            StringBuilder sb = new StringBuilder();
            sb.Append("<file>");
            sb.AppendFormat("<name>{0}</name>", this.FileName);
            sb.AppendFormat("<content>{0}</content>", strFile);
            sb.Append("</file>");
            this.FileData = sb.ToString();

            return this;
        }

       

        /// 

        ///Andy ZHANG
        ///@2017-7-17
        ///Update file in google drive
        [Attributes.Method("Update", MethodType.Execute, "Update File", "Update File",
                    new string[] { "FileId", "FileData" }, //required property array. Properties must match the names of Properties decorated with the Property Attribute
                    new string[] { "FileId", "FileData" }, //input property array. Properties must match the names of Properties decorated with the Property Attribute
                    new string[] { "FileId" })] //return property array. Properties must match the names of Properties decorated with the Property Attribute
        public GoogleDriveFileService UpdateFile()
        {
                var fileId = this.FileId;                       
                string credPath = this.ServiceConfiguration[credName].ToString();
                var _service = GoogleDriveHelper.GetDriveService(credPath);

                XDocument xDoc = new XDocument();
                xDoc = XDocument.Parse(this.FileData);
                string strFileName = xDoc.Root.Element("name").Value;
                byte[] bFile = Convert.FromBase64String(xDoc.Root.Element("content").Value);
                string strFileExt = System.IO.Path.GetExtension(strFileName);
                string strMimeType = GetMimeType(strFileName);
                Google.Apis.Drive.v3.Data.File file = new Google.Apis.Drive.v3.Data.File();

                // File's new metadata.
                file.MimeType = strMimeType;

                // File's new content.
                System.IO.MemoryStream stream = new System.IO.MemoryStream(bFile);
                // Send the request to the API.
                Google.Apis.Drive.v3.FilesResource.UpdateMediaUpload request = _service.Files.Update(file, fileId, stream, strMimeType);
                //request.NewRevision = newRevision;
                request.Upload();
                this.FileId = fileId;
                return this;
            
    }
        ///Andy ZHANG
        ///@2017-7-17
        ///Create Folder in google drive
        [Attributes.Method("Create Folder", MethodType.Execute, "Create Folder", "Create Folder",
                    new string[] { "FolderName"}, //required property array. Properties must match the names of Properties decorated with the Property Attribute
                    new string[] { "FolderName" }, //input property array. Properties must match the names of Properties decorated with the Property Attribute
                    new string[] { "FolderPath" })] //return property array. Properties must match the names of Properties decorated with the Property Attribute
        public GoogleDriveFileService CreateFolder()
        {
            string credPath = this.ServiceConfiguration[credName].ToString();
            var _service = GoogleDriveHelper.GetDriveService(credPath);
            string foldername = this.FolderName;
            var fileMetadata = new Google.Apis.Drive.v3.Data.File()
            {
                Name =foldername,
                MimeType = "application/vnd.google-apps.folder"
            };
            var request = _service.Files.Create(fileMetadata);
            request.Fields = "id";
            var file = request.Execute();
            this.FolderPath = file.Id;
            return this;
        }

        ///Andy ZHANG
        ///@2017-7-17
        ///Update file in google drive
        [Attributes.Method("Delete", MethodType.Delete, "Delete", "Delete",
                    new string[] { "FileId" }, //required property array. Properties must match the names of Properties decorated with the Property Attribute
                    new string[] { "FileId" }, //input property array. Properties must match the names of Properties decorated with the Property Attribute
                    new string[] { "FileId" })] //return property array. Properties must match the names of Properties decorated with the Property Attribute
        public GoogleDriveFileService DeleteFile()
        {
            string credPath = this.ServiceConfiguration[credName].ToString();
            var _service = GoogleDriveHelper.GetDriveService(credPath);
            string fileId = this.FileId;
            try
            {
                _service.Files.Delete(fileId).Execute();
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
            }
            this.FileId = fileId;
            return this;
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

        //[Attributes.Method("Delete", MethodType.Execute, "Delete a file", "Delete a file",
        //    new string[] { "FolderPath", "FileName" }, //required property array. Properties must match the names of Properties decorated with the Property Attribute
        //    new string[] { "FolderPath", "FileName" }, //input property array. Properties must match the names of Properties decorated with the Property Attribute
        //    new string[]{})] //return property array. Properties must match the names of Properties decorated with the Property Attribute
        //public GoogleDriveFileService Delete()
        //{
        //    string sBasePath = this.ServiceConfiguration["File Path"].ToString();
        //    // append a slash if it is missing
        //    if (!sBasePath.EndsWith(@"\"))
        //    {
        //        sBasePath += @"\";
        //    }

        //    string sFolderPath = this.FolderPath;
        //    if (!sFolderPath.EndsWith(@"\"))
        //    {
        //        sFolderPath += @"\";
        //    }

        //    string sFolderFullPath = sBasePath + sFolderPath;
        //    SIO.File.Delete(sFolderFullPath + this.FileName);
        //    return this;
        //}
        #endregion
    }
}
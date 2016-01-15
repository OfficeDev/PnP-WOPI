using com.microsoft.dx.officewopi.Models;
using Microsoft.WindowsAzure;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace com.microsoft.dx.officewopi.Utils
{
    public class AzureStorageUtil
    {
        public static async Task<List<FileModel>> GetFiles(string container)
        {
            //initialize return
            List<FileModel> files = new List<FileModel>();

            //get configuration data
            string azureBlobProtocol = ConfigurationManager.AppSettings["abs:Protocol"];
            string azureBlobAccountName = ConfigurationManager.AppSettings["abs:AccountName"];
            string azureBlobAccountkey = ConfigurationManager.AppSettings["abs:AccountKey"];
            container = container.Replace(".", "");

            // Initialize the Azure account information
            string connString = string.Format("DefaultEndpointsProtocol={0};AccountName={1};AccountKey={2}",
                azureBlobProtocol, azureBlobAccountName, azureBlobAccountkey);
            CloudStorageAccount cloudStorageAccount = CloudStorageAccount.Parse(connString);

            // Create the blob client, which provides authenticated access to the Blob service.
            CloudBlobClient blobClient = cloudStorageAccount.CreateCloudBlobClient();

            // Get the container reference...create if it does not exist
            CloudBlobContainer blobContainer = blobClient.GetContainerReference(container);
            if (!blobContainer.Exists())
                await blobContainer.CreateAsync();
            else
            {
                //get blobs from container
                var blobs = blobContainer.ListBlobs(null, true, BlobListingDetails.All);
                foreach (CloudBlockBlob blob in blobs)
                {
                    files.Add(new FileModel() { BaseFileName = blob.Name, Size = blob.Properties.Length });
                }
            }
            return files;
        }


        public static async Task<string> UploadFile(string fileName, string container, byte[] fileBytes)
        {
            //get configuration data
            string url = null;
            string azureBlobProtocol = ConfigurationManager.AppSettings["abs:Protocol"];
            string azureBlobAccountName = ConfigurationManager.AppSettings["abs:AccountName"];
            string azureBlobAccountkey = ConfigurationManager.AppSettings["abs:AccountKey"];
            container = container.Replace(".", "");

            // Initialize the Azure account information
            string connString = string.Format("DefaultEndpointsProtocol={0};AccountName={1};AccountKey={2}",
                azureBlobProtocol, azureBlobAccountName, azureBlobAccountkey);
            CloudStorageAccount cloudStorageAccount = CloudStorageAccount.Parse(connString);

            // Create the blob client, which provides authenticated access to the Blob service.
            CloudBlobClient blobClient = cloudStorageAccount.CreateCloudBlobClient();

            // Get the container reference...create if it does not exist
            CloudBlobContainer blobContainer = blobClient.GetContainerReference(container);
            if (!blobContainer.Exists())
                await blobContainer.CreateAsync();
            CloudBlockBlob blob = blobContainer.GetBlockBlobReference(fileName);

            // Set permissions on the container.
            BlobContainerPermissions containerPermissions = new BlobContainerPermissions();
            containerPermissions.PublicAccess = BlobContainerPublicAccessType.Off;
            await blobContainer.SetPermissionsAsync(containerPermissions);

            //upload the file using a memory stream
            using (MemoryStream stream = new MemoryStream(fileBytes))
            {
                stream.Write(fileBytes, 0, fileBytes.Length);
                stream.Seek(0, SeekOrigin.Begin);
                stream.Flush();
                blob.UploadFromStream(stream);
                stream.Close();

                //get the url of the blob
                url = String.Format("{0}://{1}.blob.core.windows.net/{2}/{3}", azureBlobProtocol,
                    azureBlobAccountName, container, fileName);
            }

            return url;
        }
        public static async Task<bool> DeleteFile(string fileName, string container)
        {
            //get configuration data
            string azureBlobProtocol = ConfigurationManager.AppSettings["abs:Protocol"];
            string azureBlobAccountName = ConfigurationManager.AppSettings["abs:AccountName"];
            string azureBlobAccountkey = ConfigurationManager.AppSettings["abs:AccountKey"];
            container = container.Replace(".", "");

            // Initialize the Azure account information
            string connString = string.Format("DefaultEndpointsProtocol={0};AccountName={1};AccountKey={2}",
                azureBlobProtocol, azureBlobAccountName, azureBlobAccountkey);
            CloudStorageAccount cloudStorageAccount = CloudStorageAccount.Parse(connString);

            // Create the blob client, which provides authenticated access to the Blob service.
            CloudBlobClient blobClient = cloudStorageAccount.CreateCloudBlobClient();

            // Get the container reference...create if it does not exist
            CloudBlobContainer blobContainer = blobClient.GetContainerReference(container);
            if (!blobContainer.Exists())
                return true;

            try
            {
                CloudBlockBlob blob = blobContainer.GetBlockBlobReference(fileName);
                await blob.DeleteAsync();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static async Task<byte[]> GetFile(string fileName, string container)
        {
            //get configuration data
            string azureBlobProtocol = ConfigurationManager.AppSettings["abs:Protocol"];
            string azureBlobAccountName = ConfigurationManager.AppSettings["abs:AccountName"];
            string azureBlobAccountkey = ConfigurationManager.AppSettings["abs:AccountKey"];

            // Initialize the Azure account information
            string connString = string.Format("DefaultEndpointsProtocol={0};AccountName={1};AccountKey={2}",
                azureBlobProtocol, azureBlobAccountName, azureBlobAccountkey);
            CloudStorageAccount cloudStorageAccount = CloudStorageAccount.Parse(connString);

            // Create the blob client, which provides authenticated access to the Blob service.
            CloudBlobClient blobClient = cloudStorageAccount.CreateCloudBlobClient();

            // Get the container reference...create if it does not exist
            CloudBlobContainer blobContainer = blobClient.GetContainerReference(container);
            CloudBlockBlob blob = blobContainer.GetBlockBlobReference(fileName);

            // Read the blob
            var b = await blob.OpenReadAsync();
            byte[] bytes = new byte[b.Length];
            b.Read(bytes, 0, (int)b.Length);
            return bytes;
        }

        /*
        public static bool FileExists(string fileName, string container)
        {
            //get configuration data
            string azureBlobProtocol = ConfigurationManager.AppSettings["abs:Protocol"];
            string azureBlobAccountName = ConfigurationManager.AppSettings["abs:AccountName"];
            string azureBlobAccountkey = ConfigurationManager.AppSettings["abs:AccountKey"];

            // Initialize the Azure account information
            string connString = string.Format("DefaultEndpointsProtocol={0};AccountName={1};AccountKey={2}",
                azureBlobProtocol, azureBlobAccountName, azureBlobAccountkey);
            CloudStorageAccount cloudStorageAccount = CloudStorageAccount.Parse(connString);

            // Create the blob client, which provides authenticated access to the Blob service.
            CloudBlobClient blobClient = cloudStorageAccount.CreateCloudBlobClient();

            // Get the container reference...create if it does not exist
            CloudBlobContainer blobContainer = blobClient.GetContainerReference(container);
            if (!blobContainer.Exists())
                blobContainer.Create();
            CloudBlockBlob blob = blobContainer.GetBlockBlobReference(fileName);

            // Set permissions on the container.
            BlobContainerPermissions containerPermissions = new BlobContainerPermissions();
            containerPermissions.PublicAccess = BlobContainerPublicAccessType.Off;
            blobContainer.SetPermissions(containerPermissions);

            if (AzureStorageUtil.Exists(blob))
                return true;
            else
                return false;
        }

        private static bool Exists(CloudBlob blob)
        {
            try
            {
                blob.FetchAttributes();
                return true;
            }
            catch (StorageClientException e)
            {
                if (e.ErrorCode == StorageErrorCode.ResourceNotFound)
                    return false;
                else
                    throw;
            }
        }
        */
    }
}
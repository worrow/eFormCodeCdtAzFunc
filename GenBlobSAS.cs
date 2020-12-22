using System;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Azure.Storage.Sas;
using Azure.Storage;
namespace GenerateBlobSASNSave2DB
{
    public static class GenBlobSAS
    {
        [FunctionName("GenBlobSAS")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("Processing a HTTP request.");


            string blobname = req.Headers["blobname"];
            string originalname = req.Headers["originalname"];
            string seq = req.Headers["videoSeq"];
            string lang = req.Headers["language"];
            string userID = req.Headers["userID"];
            string description = String.Empty;
            if (lang == "TC")
            {
                if (Int32.Parse(seq) == 1)
                    description = "繁體中文視頻一";
                else if (Int32.Parse(seq) == 2)
                    description = "繁體中文視頻二";
                else
                    description = "繁體中文視頻三";
            }
            else if (lang == "SC")
            {
                if (Int32.Parse(seq) == 1)
                    description = "简体中文视频一";
                else if (Int32.Parse(seq) == 2)
                    description = "简体中文视频二";
                else
                    description = "简体中文视频三";
            }
            else
            {
                if (Int32.Parse(seq) == 1)
                    description = "English Video 1";
                else if (Int32.Parse(seq) == 2)
                    description = "English Video 2";
                else
                    description = "English Video 3";
            }

            string responseMessage = String.Empty;
           
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            
            string sasURI = GetContainerSasUri(blobname);

            try
            {
                using (SqlConnection connection = new SqlConnection(System.Environment.GetEnvironmentVariable("SQLDB"))) // "Server=tcp:hkdc03doutdchdb.database.windows.net,1433;Initial Catalog=hkdc03d_DCHePortalDEV;Persist Security Info=False;User ID=osadmin-sql;Password=mQ7nUM-j6Du9y;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.Connection = connection;
                        cmd.CommandText = "sp_CodeCdt_UploadVideo";
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@sasURI", sasURI);
                        cmd.Parameters.AddWithValue("@originalName", originalname);
                        cmd.Parameters.AddWithValue("@description", description);
                        cmd.Parameters.AddWithValue("@lang", lang);
                        cmd.Parameters.AddWithValue("@seq", seq);
                        cmd.Parameters.AddWithValue("@userID", userID);
                        connection.Open();
                        cmd.ExecuteNonQuery();
                    }
                }
                log.LogInformation("Update one record at " + DateTime.Now.ToLongTimeString());
                responseMessage = "File uploaded successfully";
            }
            catch (Exception ex)
            {
                log.LogError(ex.Message);
                responseMessage = ex.Message;
            }

            return new OkObjectResult(responseMessage);
        }


        private static string GetContainerSasUri(string blobname)
        {
            string StorageAccountName = System.Environment.GetEnvironmentVariable("StorageAccountName");
            string StorageAccountKey = System.Environment.GetEnvironmentVariable("StorageAccountKey");

            AccountSasBuilder sas = new AccountSasBuilder
            {
                // Allow access to blobs
                Services = AccountSasServices.Blobs,

                // Allow access to the service level APIs
                ResourceTypes = AccountSasResourceTypes.All,
                StartsOn = DateTimeOffset.UtcNow,
                // Access expires in 1 hour!
                ExpiresOn = DateTimeOffset.UtcNow.AddDays(99999)
            };
            sas.Protocol = SasProtocol.Https;


            sas.SetPermissions(AccountSasPermissions.Read);

            // Create a SharedKeyCredential that we can use to sign the SAS token
            StorageSharedKeyCredential credential = new StorageSharedKeyCredential(StorageAccountName, StorageAccountKey);
            UriBuilder sasUri = new UriBuilder(System.Environment.GetEnvironmentVariable("BlobURI") + blobname);
            sasUri.Query = sas.ToSasQueryParameters(credential).ToString();
            return sasUri.Uri.AbsoluteUri;
        }
    }
}

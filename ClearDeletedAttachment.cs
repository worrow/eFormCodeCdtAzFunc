using System;
using System.Collections.Generic;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using System.Data;
using System.Data.SqlClient;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;

namespace GenerateBlobSASNSave2DB
{
    internal class ClearDeletedAttachment
    {
        private class BlobInfo
        {
            public int id { get; set; }
            public string BlobFileName { get; set; }
            public DateTime CreatedDate { get; set; }
        }
        [FunctionName("ClearDeletedAttachment")]
        public static void Run([TimerTrigger("0 0 */ * *")] TimerInfo myTimer, ILogger log)
        {//"0 0 */ * *"
            List<BlobInfo> lst = new List<BlobInfo>();
            BlobInfo item;
            try
            {
                using (SqlConnection conn = new SqlConnection(Environment.GetEnvironmentVariable("SQLDB")))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "sp_CodeCdt_GetBlobDeleteList";
                        cmd.CommandType = CommandType.StoredProcedure;

                        conn.Open();
                        SqlDataReader dr = cmd.ExecuteReader();
                        while (dr.Read())
                        {
                            item = new BlobInfo();
                            item.id = dr.GetInt32(0);
                            item.BlobFileName = dr.GetString(1);
                            item.CreatedDate = dr.GetDateTime(2);
                            lst.Add(item);
                        }
                    }

                }
                //if(lst.Count > 0)
                //    latestBlobCreatedDate = lst.Max(x => x.CreatedDate);


                BlobServiceClient blobServiceClient = new BlobServiceClient(Environment.GetEnvironmentVariable("AzureWebJobsStorage"));

                BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient("eportal-codecdt");

                using (SqlConnection conn = new SqlConnection(Environment.GetEnvironmentVariable("SQLDB")))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "sp_CodeCdt_ClearBlobDeleteList";
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id", SqlDbType.Int);
                        conn.Open();
                        foreach (BlobInfo s in lst)
                        {
                            cmd.Parameters[0].Value = s.id;
                            containerClient.DeleteBlobIfExists("CodeCdtAttachments/" + s.BlobFileName, DeleteSnapshotsOption.IncludeSnapshots);
                            cmd.ExecuteNonQuery();
                        }
                    }
                }

               
            }
            catch (Exception ex)
            {
                log.LogError(ex.Message);
            }
        }
    }
}

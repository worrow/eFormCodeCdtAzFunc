using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;

namespace GenerateBlobSASNSave2DB
{
	public static class BlobDBUpdate
	{
		[FunctionName("BlobDBUpdate")]
		public static void Run([BlobTrigger("eportal-codecdt/{name}", Connection = "AzureWebJobsStorage")] Stream myBlob, string name, ILogger log)
		{
			string codeCdtID = string.Empty;
			string uploadUser = string.Empty;
			string orignalFileName = string.Empty;
			int ccID = 0;
		//	log.LogInformation($"C# Blob trigger function Processed blob\n Name:{name} \n Size: {myBlob.Length} Bytes");
			try
			{
				//log.LogInformation("$$$$$$$$$$$$$$$$" + name + "$$$$$$$$$$$$$$$$");
				if (!name.StartsWith("CodeCdtAttachments"))
				{
					return;
				}
				codeCdtID = name.Split('_')[1];
				uploadUser = name.Split('_')[2];
				orignalFileName = name.Substring(name.IndexOf(name.Split('_')[3]));
				//log.LogInformation("**************orignalFileName**************: " + orignalFileName);
				if (!int.TryParse(codeCdtID, out ccID))
				{
					throw new Exception("Incorrect or obsolete Filename Format");
				}
				//log.LogInformation("##############BlobName" + name.Replace("CodeCdtAttachments", "") + "##########################");
				using SqlConnection conn = new SqlConnection(Environment.GetEnvironmentVariable("SQLDB"));
				//log.LogInformation("*******************************************");
				//log.LogInformation("CCCCCCCCCCCCCCCCCCId is " + codeCdtID);
				using SqlCommand cmd = new SqlCommand();
				cmd.Connection = conn;
				cmd.CommandText = "sp_CodeCdt_Attachment";
				cmd.CommandType = CommandType.StoredProcedure;
				cmd.Parameters.AddWithValue("@codeCdtID", codeCdtID);
				cmd.Parameters.AddWithValue("@uploadUser", uploadUser);
				cmd.Parameters.AddWithValue("@OriginalFileName", orignalFileName);
				cmd.Parameters.AddWithValue("@BlobFileName", name.Replace("CodeCdtAttachments/", ""));
				conn.Open();
				cmd.ExecuteNonQuery();
			}
			catch (Exception ex)
			{
				log.LogError(ex.Message);
				throw;
			}
		}
	}

}

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
using ExcelDataReader;
using System.Web.Http;

namespace GenerateBlobSASNSave2DB
{
    public static class ImportCodeCdt
    {
        [FunctionName("ImportCodeCdt")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            IExcelDataReader excelReader = null;
            MemoryStream ms = new MemoryStream();
            await req.Body.CopyToAsync(ms);
            try
            {
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(ms);
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };
                DataSet dataSet = excelReader.AsDataSet(conf);
                 using (SqlConnection conn = new SqlConnection(System.Environment.GetEnvironmentVariable("SQLDB")))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.CommandText = "sp_CodeCdt_Import";
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Connection = conn;
                        cmd.CommandTimeout = 9999;
                        SqlParameter sqlParam = cmd.Parameters.AddWithValue("@ImportTable", dataSet.Tables[0]);
                        sqlParam.SqlDbType = SqlDbType.Structured;
                        conn.Open();
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                log.LogError(ex.Message);
                return new InternalServerErrorResult();
            }
            //string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            // parse query parameter
            

            return new OkObjectResult("Import Success");
        }
    }
}

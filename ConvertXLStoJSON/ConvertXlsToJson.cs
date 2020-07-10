using ExcelDataReader;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Nancy.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;

namespace ConvertXLStoJSON
{
	public static class ConvertXlsToJson
    {
        [FunctionName("ConvertXlsToJson")]
		public static IActionResult Run(
			[HttpTrigger(AuthorizationLevel.Function, "post", "put", "patch", Route = null)] HttpRequest req,
			ILogger log)
		{
			Stream blob = req.Body;
			DataSet ds = CreateDataSet(blob);
			foreach (DataTable table in ds.Tables)
			{
				return (ActionResult)new OkObjectResult(DataTableToJSON(table));
			}

			return blob != null
			? (ActionResult)new OkObjectResult("")
			: new BadRequestObjectResult("Please pass a blob in the request body");
		}

		public static string DataTableToJSON(DataTable table)
        {
            List<Dictionary<string, object>> list = new List<Dictionary<string, object>>();

            foreach (DataRow row in table.Rows)
            {
                Dictionary<string, object> dict = new Dictionary<string, object>();

                foreach (DataColumn col in table.Columns)
                {
                    dict[col.ColumnName] = (Convert.ToString(row[col]));
                }
                list.Add(dict);
            }
            JavaScriptSerializer serializer = new JavaScriptSerializer();

            return serializer.Serialize(list);
        }

        public static DataSet CreateDataSet(Stream stream)
        {
            DataSet ds = null;
           
			try
			{
				System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
				IExcelDataReader reader = null;
				reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
				ds = reader.AsDataSet(new ExcelDataSetConfiguration()
				{
					ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
					{
						UseHeaderRow = true,
					}
				});
			}
			catch (System.Exception ex)
			{
				if (ex.Message == "You can ignore bad data by setting BadDataFound to null.")
				{
					using (ZipArchive archive = new ZipArchive(stream))
					{
						foreach (ZipArchiveEntry entry in archive.Entries)
						{
							using (var fileStream = entry.Open())
							{
								ds = CreateDataSet(fileStream);
							}
						}
					}
				}
				else
				{
					throw ex;
				}
			}

			return ds;
        }
    }
}






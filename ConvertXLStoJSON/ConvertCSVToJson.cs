using CsvHelper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;

namespace ConvertXLStoJSON
{
	public static class ConvertCSVToJson
    {
		[FunctionName("ConvertCSVToJson")]
		public static IActionResult Run(
			[HttpTrigger(AuthorizationLevel.Function, "post", "put", "patch", Route = null)] HttpRequest req,
			ILogger log)
		{
			Stream blob = req.Body;
			return (ActionResult)new OkObjectResult(Convert(blob));
		}

		public static string Convert(Stream blob)
        {
			var returnValue = "";

			try
			{
				var sReader = new StreamReader(blob);
				var csv = new CsvReader(sReader, CultureInfo.InvariantCulture);

				csv.Read();
				csv.ReadHeader();

				var csvRecords = csv.GetRecords<object>().ToList();
				
				returnValue = JsonConvert.SerializeObject(csvRecords);
			}
			catch (System.Exception ex)
			{
				if (ex.Message == "You can ignore bad data by setting BadDataFound to null.")
				{
					using (ZipArchive archive = new ZipArchive(blob))
					{
						foreach (ZipArchiveEntry entry in archive.Entries)
						{
							using (var fileStream = entry.Open())
							{
								returnValue = Convert(fileStream);
							}
						}
					}
				}
				else
				{
					throw ex;
				}		
			}

            return returnValue;
        }
    }
}

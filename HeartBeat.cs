using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;

namespace GenerateBlobSASNSave2DB
{
	internal class HeartBeat
	{
		[FunctionName("HeartBeat")]
		public static void Run([TimerTrigger("0 */14 * * * *")] TimerInfo myTimer, ILogger log)
		{//0 0/15 9-22 ? * 1/1 *
			log.LogInformation($"The heart is beating at: {DateTime.Now}");
		}
	}

}

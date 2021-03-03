using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace docx_to_json
{
    public static class fnConvertFromBlob
    {
        [FunctionName("fnConvertFromBlob-RunFromBlob")]
        public static void RunFromBlob([BlobTrigger("input-docx/{name}", Connection = "ConnectionStrings:DocxBlobConnectionString")]Stream docxBlob,
            [Blob("input-json/{name}.json", FileAccess.Write, Connection = "ConnectionStrings:DocxBlobConnectionString")] Stream jsonBlob, 
            string name, ILogger log)
        {
            // start tracing
            var traceID = Guid.NewGuid().ToString();
            log.LogInformation($"{traceID} - fnConvertFromBlob-RunFromBlob - blob\n Name:{name} \n Size: {docxBlob.Length} Bytes");

            // convert the file
            var converter = new DocxToJsonConvertor(traceID, log);
            converter.StringsToRemove = new string[] { "\r\n", "<sup>", "</sup>", "&#x200e;", "<br />" };
            var jsonOutput = converter.ConvertDocx(docxBlob);

            // write output file
            using (var writer = new StreamWriter(jsonBlob))
            {
                writer.Write(jsonOutput);
                writer.Flush();
            }
        }

        [FunctionName("fnConvertFromBlob-RunFromHttp")]
        public static async Task<IActionResult> RunFromHttp([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            // start tracing
            var traceID = Guid.NewGuid().ToString();
            log.LogInformation($"{traceID} - fnConvertFromBlob-RunFromHttp");

            // validate the request includes the blob URL
            string uriString = req.Query["blobUri"];
            string requestBody = String.Empty;
            using (StreamReader streamReader = new StreamReader(req.Body))
            {
                requestBody = await streamReader.ReadToEndAsync();
            }
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            uriString = uriString ?? data?.blobUri;
            Uri blobUri;
            if (uriString == null || Uri.TryCreate(uriString, UriKind.Absolute, out blobUri))
            {
                return new BadRequestObjectResult("Please supply a valid URL to the blob on the query string or in the request body, with a name of 'blobUri'");
            }

            // convert the file
            var converter = new DocxToJsonConvertor(traceID, log);
            converter.StringsToRemove = new string[] { "\r\n", "<sup>", "</sup>", "&#x200e;", "<br />" };
            var jsonOutput = converter.ConvertDocx(blobUri);

            // return the result
            return new OkObjectResult(jsonOutput);
        }

    }
}

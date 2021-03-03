using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace docx_to_json
{
    public static class fnConvertFromHttp
    {
        [FunctionName("fnConvertFromHttp")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            // start tracing
            var traceID = Guid.NewGuid().ToString();
            log.LogInformation($"{traceID} - fnConvertFromHttp");

            // validate body
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            string base64data = data?.base64data;
            if (String.IsNullOrEmpty(base64data))
            {
                return new BadRequestObjectResult("Payload is expected in the format: { \"base64data\": \"abcd...\" }");
            }

            // convert the base 64 data into a stream and convert to JSON
            var jsonResult = new JObject();
            var bytes = Convert.FromBase64String(base64data);
            using (var fileStream = new MemoryStream(bytes))
            {
                var converter = new DocxToJsonConvertor(traceID, log);
                converter.StringsToRemove = new string[] { "\r\n", "<sup>", "</sup>", "&#x200e;", "<br />" };
                jsonResult = converter.ConvertDocx(fileStream);
            }

            // return the json result
            return new OkObjectResult(jsonResult);
        }
    }
}

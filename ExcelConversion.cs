using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Azure.Storage.Blobs;
using NPOI.SS.UserModel;



namespace AzureFormRecognizer.Preparation
{
    public static class AFRPreparation
    {
        [FunctionName("ExcelPrintSetup")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");
            // Use storage account connection string from app settings
            string connectionString = Environment.GetEnvironmentVariable("StorageAccountConnectionString");

            // Check for container name in query or use environmental variable
            string srcContainerName = req.Query["sourceContainer"];
            srcContainerName = srcContainerName ?? Environment.GetEnvironmentVariable("SourceBlobContainer");
            
            string destContainerName = req.Query["destinationContainer"];
            destContainerName = destContainerName ?? Environment.GetEnvironmentVariable("DestinationBlobContainer");

            // Get blob filepath (does not include container)
            string name = req.Query["blobName"];

            string responseMessage;
            try {
                // Connect to blob storage
                BlobContainerClient srcContainer = new BlobContainerClient(connectionString, srcContainerName);
                var blob = srcContainer.GetBlobClient(name);

                IWorkbook wb;
                using (MemoryStream memoryStream = new MemoryStream()) {
                    await blob.DownloadToAsync(memoryStream);
                    memoryStream.Position = 0;
                    // Open workbook using WorkbookFactory which accepts xlsx and xls.
                    wb = WorkbookFactory.Create(memoryStream);
                }

                // For each sheet, set margins, landscape orientation, and fit to column width
                int numSheets = wb.NumberOfSheets;
                for (int n = 0; n < numSheets; n++) {
                    ISheet sheet = wb.GetSheetAt(n);
                    sheet.SetMargin(MarginType.RightMargin, 0.05d);
                    sheet.SetMargin(MarginType.TopMargin, 0.05d);
                    sheet.SetMargin(MarginType.LeftMargin, 0.05d);
                    sheet.SetMargin(MarginType.BottomMargin, 0.05d);
                    sheet.PrintSetup.Landscape = true;
                    sheet.FitToPage = true; // Enables auto fit columns
                    sheet.PrintSetup.FitWidth = 1; // 
                    sheet.PrintSetup.FitHeight = 0; // 
                }
                
                BlobContainerClient destContainer = new BlobContainerClient(connectionString, destContainerName);
                var destBlob = destContainer.GetBlobClient(name);

                using (MemoryStream outputMemoryStream = new MemoryStream()) {
                    wb.Write(outputMemoryStream, true);
                    outputMemoryStream.Position = 0;
                    await destBlob.UploadAsync(outputMemoryStream, true);
                }
                
                responseMessage = $"File was successfully processed";

            } catch (Exception e) {
                responseMessage = $"File was not successfully processed. Exception message: {e.Message}";
            }


            return new OkObjectResult(responseMessage);
        }
    }

}

# Excel Print Setup Function

This repository is for an Azure Function App that prepares Excel spreadsheet files (.xls and .xlsx) for automated PDF creation.
When converting Excel spreadsheets to PDF, the output using the default settings often suffers from formatting and layout errors including split horizontal tables, poor font scaling, and undesirable page order. 
These formatting problems often make the output incompatible with data extraction tools like [Azure Form Recognizer](https://learn.microsoft.com/en-us/azure/applied-ai-services/form-recognizer/?view=form-recog-3.0.0) where a logical document layout is critical for a successful extraction.

## Azure Function

The Azure Function App utilizes the .NET [NPOI](https://github.com/nissl-lab/npoi) library (which is the .NET version of Apache's [POI](https://poi.apache.org/) project) to edit the page and print configurations of Excel workbooks.
This allows editing Excel workbooks within a consumption-based Function App without need of external applications.

The configuration changes to improve PDF export formatting are:
- Reduction of print margins to 0.05"
- Conversion to landscape orientation
- Full-width scaling (i.e. "Fit All Columns on One Page")

This Azure Function is intended for use within a complete automation solution leveraging an Azure Logic App. A simplified infrastructure diagram of such a solution is shown below.

![Infrastructure Map](/images/infra.png)


## Invoice Example

Example business invoice document using an Excel template:
![Example Excel Invoice](/images/ExcelInvoice.png)

Output of a PDF conversion without `ExcelPrintSetup`
![Sample default output](/images/defaultoutput.png)

Output of a PDF conversion after using `ExcelPrintSetup` function
![Sample processed output](/images/processedoutput.png)

## Deployment

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Fbwilliams2%2FExcelPrintSetupFunction%2Fmain%2Fazuredeploy.json)

After resource creation using the "Deploy to Azure" button above, deploy the Azure Functions using your preferred [deployment method](https://learn.microsoft.com/en-us/azure/azure-functions/functions-deployment-technologies).
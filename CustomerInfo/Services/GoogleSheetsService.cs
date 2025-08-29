using ClosedXML.Excel;
using CustomerInfo.Model;
using DocumentFormat.OpenXml.Presentation;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System.Text;

namespace CustomerInfo.Services
{
    public class GoogleSheetsService : IExcelExportService
    {
       
        private readonly SheetsService _sheetsService;
        private readonly string _spreadsheetId;
        private const string SheetName = "CustomerData";
        private readonly string _filePath;

        public GoogleSheetsService(IConfiguration config)
        {
            var googleCredentialsJson = Environment.GetEnvironmentVariable("GOOGLE_CREDENTIALS_JSON");
            GoogleCredential credential;
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(googleCredentialsJson)))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(SheetsService.Scope.Spreadsheets);
            }

            var sheetsService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "Company Registration Form"
            });
            var spreadsheetId = Environment.GetEnvironmentVariable("GOOGLE_SPREADSHEET_ID") ?? "your-spreadsheet-id";

            var valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>> { new List<object> { "Full Name", "Phone", "Position", "Company" } };
            var appendRequest = sheetsService.Spreadsheets.Values.Append(valueRange, spreadsheetId, "CustomerData!A:D");
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            appendRequest.Execute();

            //var credentialsPath = Environment.GetEnvironmentVariable("GOOGLE_CREDENTIALS_PATH")?? config["GoogleSheets:CredentialsPath"];
            //var spreadsheetId = Environment.GetEnvironmentVariable("GOOGLE_SPREADSHEET_ID")?? config["GoogleSheets:SpreadsheetId"];
            //GoogleCredential credential;
            //using (var stream = new FileStream(credentialsPath, FileMode.Open, FileAccess.Read))
            //{
            //    credential = GoogleCredential.FromStream(stream)
            //        .CreateScoped(SheetsService.Scope.Spreadsheets);
            //}

            //_sheetsService = new SheetsService(new BaseClientService.Initializer()
            //{
            //    HttpClientInitializer = credential,
            //    ApplicationName = "Company Registration Form"
            //});


            //_spreadsheetId = spreadsheetId;
        }

        //public void AddRecord(FormSubmission form)
        //{
        //    var range = $"{SheetName}!A:D"; 
        //    var valueRange = new ValueRange();

        //    var objectList = new List<object>
        //    {
        //        form.FullName,
        //        form.PhoneNumber,
        //        form.Position,
        //        form.Company
        //    };

        //    valueRange.Values = new List<IList<object>> { objectList };

        //    var appendRequest = _sheetsService.Spreadsheets.Values.Append(valueRange, _spreadsheetId, range);
        //    appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
        //    appendRequest.Execute();
        //}




        //public void AddRecord(FormSubmission form)
        //{
        //    // إذا الملف مش موجود نعمل واحد جديد
        //    using var workbook = System.IO.File.Exists(_filePath)
        //        ? new XLWorkbook(_filePath)
        //        : new XLWorkbook();

        //    var worksheet = workbook.Worksheets.Any()
        //        ? workbook.Worksheet(1)
        //        : workbook.AddWorksheet("Submissions");

        //    int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        //    int newRow = lastRow + 1;

        //    if (lastRow == 0)
        //    {
        //        worksheet.Cell(1, 1).Value = "Full Name";
        //        worksheet.Cell(1, 2).Value = "Phone Number";
        //        worksheet.Cell(1, 3).Value = "Position";
        //        worksheet.Cell(1, 4).Value = "Company";
        //        newRow = 2;
        //    }

        //    worksheet.Cell(newRow, 1).Value = form.FullName;
        //    worksheet.Cell(newRow, 2).Value = form.PhoneNumber;
        //    worksheet.Cell(newRow, 3).Value = form.Position;
        //    worksheet.Cell(newRow, 4).Value = form.Company;

        //    workbook.SaveAs(_filePath);
        //}
    }


}


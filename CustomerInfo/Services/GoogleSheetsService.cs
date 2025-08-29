using CustomerInfo.Model;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace CustomerInfo.Services
{
    public class GoogleSheetsService : IExcelExportService
    {
       
        private readonly SheetsService _sheetsService;
        private readonly string _spreadsheetId;
        private const string SheetName = "CustomerData";
        public GoogleSheetsService(IConfiguration config)
        {
            // نقرأ مسار ملف الـ JSON من appsettings.json
            var credentialsPath = config["GoogleSheets:CredentialsPath"];

            GoogleCredential credential;
            using (var stream = new FileStream(credentialsPath, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(SheetsService.Scope.Spreadsheets);
            }

            _sheetsService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "Company Registration Form"
            });

            // ID الجدول (من رابط Google Sheet)
            _spreadsheetId = config["GoogleSheets:SpreadsheetId"];
        }

        public void AddRecord(FormSubmission form)
        {
            var range = $"{SheetName}!A:D"; // الأعمدة A-D
            var valueRange = new ValueRange();

            var objectList = new List<object>
            {
                form.FullName,
                form.PhoneNumber,
                form.Position,
                form.Company
            };

            valueRange.Values = new List<IList<object>> { objectList };

            var appendRequest = _sheetsService.Spreadsheets.Values.Append(valueRange, _spreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            appendRequest.Execute();
        }

       
    }


}


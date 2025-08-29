using Amazon.S3;
using Amazon.S3.Model;
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


        private readonly string accessKey;
        private readonly string secretKey;
        private readonly string region;


        public GoogleSheetsService(IConfiguration config)
        {
            accessKey = Environment.GetEnvironmentVariable("AWS_ACCESS_KEY_ID");
            secretKey = Environment.GetEnvironmentVariable("AWS_SECRET_ACCESS_KEY");
          
        }







        public async Task AddRecordAsync(FormSubmission form)
        {
            using var s3Client = new AmazonS3Client(accessKey, secretKey, Amazon.RegionEndpoint.EUCentral1);

            var getRequest = new GetObjectRequest
            {
                BucketName = "kiosk-iih-v2",
                Key = "registrations.xlsx"
            };

            using var response = await s3Client.GetObjectAsync(getRequest);
            using var memoryStream = new MemoryStream();
            await response.ResponseStream.CopyToAsync(memoryStream);

            memoryStream.Position = 0; // مهم جداً قبل فتح الـ stream
            XLWorkbook workbook;
            try
            {
                workbook = new XLWorkbook(memoryStream);
            }
            catch
            {
                // إذا الملف غير موجود أو فارغ، نعمل واحد جديد
                workbook = new XLWorkbook();
            }

            var worksheet = workbook.Worksheets.Any() ? workbook.Worksheet(1) : workbook.AddWorksheet("Submissions");

            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
            int newRow = lastRow + 1;

            if (lastRow == 0)
            {
                worksheet.Cell(1, 1).Value = "Full Name";
                worksheet.Cell(1, 2).Value = "Phone Number";
                worksheet.Cell(1, 3).Value = "Position";
                worksheet.Cell(1, 4).Value = "Company";
                newRow = 2;
            }

            worksheet.Cell(newRow, 1).Value = form.FullName;
            worksheet.Cell(newRow, 2).Value = form.PhoneNumber;
            worksheet.Cell(newRow, 3).Value = form.Position;
            worksheet.Cell(newRow, 4).Value = form.Company;

            // حفظ التغييرات على S3
            using var uploadStream = new MemoryStream();
            workbook.SaveAs(uploadStream);
            uploadStream.Position = 0;

            var putRequest = new PutObjectRequest
            {
                BucketName = "kiosk-iih-v2",
                Key = "registrations.xlsx",
                InputStream = uploadStream
            };

            await s3Client.PutObjectAsync(putRequest);
        }
    }


}


using CustomerInfo.Model;

namespace CustomerInfo.Services
{
    public interface IExcelExportService
    {
        //byte[] GenerateExcel(FormSubmission submission);
     
        Task AddRecordAsync(FormSubmission form);
    }
}

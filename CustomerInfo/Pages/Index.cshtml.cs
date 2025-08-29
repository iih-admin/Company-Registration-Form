using CustomerInfo.Model;
using CustomerInfo.Services;
using DocumentFormat.OpenXml.ExtendedProperties;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace CustomerInfo.Pages
{
    public class IndexModel : PageModel
    {
        private readonly IExcelExportService _excelExportService;
        [BindProperty]
        public FormSubmission Form { get; set; }
        public IndexModel(IExcelExportService excelExportService)
        {
            
            _excelExportService = excelExportService;
        }

        public void OnGet()
        {

        }
        public IActionResult OnPost()
        {
            if (!ModelState.IsValid) return Page();

            _excelExportService.AddRecordAsync(Form);
            //_excelExportService.GenerateExcel(Form);
            TempData["Success"] = "Your information has been saved!";
            return RedirectToPage();
        }
    }
}

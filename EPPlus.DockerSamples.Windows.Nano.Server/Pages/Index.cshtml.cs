using EPPlus.DockerSamples.Alpine;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace EPPlus.DockerSamples.Pages
{
    public class IndexModel : PageModel
    {
        private const string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        private readonly ILogger<IndexModel> _logger;

        public IndexModel(ILogger<IndexModel> logger)
        {
            _logger = logger;
        }

        public void OnGet()
        {

        }
        public FileResult OnPostCreateReport()
        {
            return File(ExcelReport.GetReport(), contentType, "EPPlusFxReport.xlsx");
        }
    }
}
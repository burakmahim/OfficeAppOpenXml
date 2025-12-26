using Microsoft.AspNetCore.Mvc;
using OfficeAppOpenXmlLibrary.Services;

namespace OfficeAppOpenXml.Core.Controllers
{
    public class SearchPageController : Controller
    {
        private readonly FileService _fileService;

        public SearchPageController(IConfiguration config)
        {
            _fileService = new FileService(config.GetConnectionString("OfficeFilesDb"));
        }

        public IActionResult Index(string? q)
        {
            if (string.IsNullOrWhiteSpace(q))
                return View(null);

            var results = _fileService.Search(q);
            return View(results);
        }
    }
}

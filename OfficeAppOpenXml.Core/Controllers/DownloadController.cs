using Microsoft.AspNetCore.Mvc;
using OfficeAppOpenXmlLibrary.Services;

namespace OfficeAppOpenXml.Core.Controllers
{
    public class DownloadController : Controller
    {
        private readonly FileService _fileService;

        public DownloadController(FileService fileService)
        {
            _fileService = fileService;
        }

        public IActionResult File(int id)
        {
            var result = _fileService.GetFile(id);
            if (result == null)
                return NotFound();

            var (fileName, mime, bytes) = result.Value;

            return File(bytes, mime, fileName);
        }


        public IActionResult Pdf(int id)
        {
            var file = _fileService.GetFileAsPdf(id);
            if (file == null) return NotFound();

            return File(file.Value.Bytes, file.Value.Mime, file.Value.FileName);
        }


    }
}

using Microsoft.AspNetCore.Mvc;
using OfficeAppOpenXmlLibrary;
using OfficeAppOpenXmlLibrary.OpenXmlFileUpload;

namespace OfficeAppOpenXml.Core.Controllers
{
    public class UploadController : Controller
    {
        [HttpGet]
        public IActionResult Index() => View();

        [HttpPost]
        [RequestSizeLimit(50_000_000)]
        public IActionResult ExtractPartial(IFormFile file, bool includeRels = true)
        {
            if (file == null || file.Length == 0)
                return PartialView("_Message", "Dosya seçilmedi.");

            using var s = file.OpenReadStream();
            var parts = OpenXmlInspector.ExtractXmlParts(s, includeRels);
            ViewBag.FileName = file.FileName;
            return PartialView("_XmlParts", parts);
        }

        [HttpPost]
        [RequestSizeLimit(50_000_000)]
        public IActionResult ConvertToDslPartial(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return PartialView("_Message", "Dosya seçilmedi.");

            using var s = file.OpenReadStream();
            var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
            string dsl = ext switch
            {
                ".pptx" => OpenXmlToDslConverter.ConvertPptxToDsl(s),
                ".xlsx" => OpenXmlToDslConverter.ConvertXlsxToDsl(s),
                _ => "Bu uzantı için DSL dönüştürücü tanımlı değil."
            };
            ViewBag.FileName = file.FileName;
            return PartialView("_Dsl", dsl);
        }

        [HttpPost]
        [RequestSizeLimit(50_000_000)]
        public IActionResult DownloadZip(IFormFile file, bool includeRels = true)
        {
            if (file == null || file.Length == 0) return BadRequest("Dosya yok.");
            using var s = file.OpenReadStream();
            var parts = OpenXmlInspector.ExtractXmlParts(s, includeRels);
            var zip = OpenXmlInspector.ToZip(parts);
            var name = Path.GetFileNameWithoutExtension(file.FileName) + "_xml.zip";
            return File(zip, "application/zip", name);
        }
    }
}

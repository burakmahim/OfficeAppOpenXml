using Microsoft.AspNetCore.Mvc;
using OfficeAppOpenXmlLibrary;                 // PowerPointLibrary / ExcelLibrary
using OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents;
using OfficeAppOpenXmlLibrary.PowerPointOpenXmlComponents;
  // IPowerPointPdfConverter, IExcelPdfConverter

namespace OfficeAppOpenXml.Core.Controllers
{
    public class PresentationController : Controller
    {


        [HttpGet]
        public IActionResult Index()
        {
            ViewBag.XmlContent = "";
            return View();
        }

        [HttpPost]
        public IActionResult DownloadPptx([FromForm] string xmlContent)
        {
            try
            {
                byte[] pptBytes = PowerPointLibrary.CreatePowerPointPresentation(xmlContent);
                return File(pptBytes,
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    "Sunum.pptx");
            }
            catch (Exception ex)
            {
                return Content("Hata oluştu: " + ex.Message);
            }
        }

        [HttpPost]
        public IActionResult GenerateExcelFromXml([FromForm] string xmlContent)
        {
            if (string.IsNullOrWhiteSpace(xmlContent))
                return BadRequest("XML boş olamaz.");

            try
            {
                byte[] result = ExcelLibrary.CreateExcel(xmlContent);
                return File(result,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "veriler.xlsx");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Excel oluşturulamadı: {ex.Message}");
            }
        }



    }
}

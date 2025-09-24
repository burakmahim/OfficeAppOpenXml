using Microsoft.AspNetCore.Mvc;
using OfficeAppOpenXmlLibrary;
using OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents;

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

        [HttpPost]
        public IActionResult ViewPdf([FromForm] string xmlContent)
        {
            try
            {
                byte[] pptBytes = PowerPointLibrary.CreatePowerPointPresentation(xmlContent);
                return File(pptBytes, "application/vnd.openxmlformats-officedocument.presentationml.presentation", "Sunum.pptx");
            }
            catch (Exception ex)
            {
                return Content("Hata oluştu: " + ex.Message);
            }
        }

        [HttpPost]
        public IActionResult GenerateExcelPdf([FromForm] string xmlContent)
        {
            try
            {
                byte[] pdfBytes = ExcelPdfConverter.ConvertToPdf(xmlContent);
                return File(pdfBytes, "application/pdf", "excel-rapor.pdf");
            }
            catch (Exception ex)
            {
                return Content("PDF dönüştürme hatası: " + ex.Message);
            }
        }
    }
}

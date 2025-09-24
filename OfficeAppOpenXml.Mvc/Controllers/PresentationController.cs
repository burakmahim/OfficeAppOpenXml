using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using OfficeAppOpenXmlLibrary;
using OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents;
using OfficeAppOpenXmlLibrary.PowerPointOpenXmlComponents;

namespace OfficeAppOpenXml.Mvc.Controllers
{
    public class PresentationController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        
        [HttpPost]
        [ValidateInput(false)]
        public ActionResult DownloadPptx(string xmlContent)
        {
            try
            {
                byte[] pptBytes = OfficeAppOpenXmlLibrary.PowerPointLibrary.CreatePowerPointPresentation(xmlContent);
                return File(pptBytes,
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation","Sunum.pptx");
            }
            catch (Exception ex)
            {
                return Content("Hata oluştu: " + ex.Message);
            }
        }

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult DownloadExcel(string xmlContent)
        {
            try
            {
                byte[] pptBytes = OfficeAppOpenXmlLibrary.ExcelLibrary.CreateExcel(xmlContent);
                return File(pptBytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","Rapor.xlsx");
            }
            catch (Exception ex)
            {
                return Content("Hata oluştu: " + ex.Message);
            }
        }

		[HttpPost]
		[ValidateInput(false)]
		public ActionResult ViewPdf(string xmlContent)
		{
			if (string.IsNullOrWhiteSpace(xmlContent))
			{
				ViewBag.Error = "XML içeriği boş gönderildi.";
				return View("Index");
			}

			try
			{
				byte[] pdfBytes = PowerPointPdfConverter.ConvertToPdf(xmlContent);
				return File(pdfBytes, "application/pdf");
			}
			catch (Exception ex)
			{
				ViewBag.Error = ex.Message;
				return View("Index");
			}
		}

		[HttpPost]
		[ValidateInput(false)]
		public ActionResult GenerateExcelPdf(string xmlContent)
		{

			if (string.IsNullOrWhiteSpace(xmlContent))
			{
				ViewBag.Error = "XML içeriği boş gönderildi.";
				return View("Index");
			}

			try
			{
				byte[] pdfBytes = ExcelPdfConverter.ConvertToPdf(xmlContent);
				return File(pdfBytes, "application/pdf");
			}

			catch (Exception ex)
			{
				ViewBag.Error = ex.Message;
				return View("Index");
			}

		}
	}
}
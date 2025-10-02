using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using OfficeAppOpenXmlLibrary;

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
                byte[] pptBytes = PowerPointLibrary.CreatePowerPointPresentation(xmlContent);
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
                byte[] pptBytes = ExcelLibrary.CreateExcel(xmlContent);
                return File(pptBytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","Rapor.xlsx");
            }
            catch (Exception ex)
            {
                return Content("Hata oluştu: " + ex.Message);
            }
        }
    }
}
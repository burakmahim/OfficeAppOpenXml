using OfficeAppOpenXmlLibrary;
using OfficeAppOpenXmlLibrary.OpenXmlFileUpload;
using System.IO;
using System.Web;
using System.Web.Mvc;

namespace OfficeAppOpenXml.Mvc.Controllers
{
    public class UploadController : Controller
    {
        [HttpGet]
        public ActionResult Index() => View();

        [HttpPost]
        public ActionResult ExtractPartial(HttpPostedFileBase file, bool includeRels = true)
        {
            if (file == null || file.ContentLength == 0)
                return PartialView("_Message", "Dosya seçilmedi.");

            var parts = OpenXmlInspector.ExtractXmlParts(file.InputStream, includeRels);
            ViewBag.FileName = file.FileName;
            return PartialView("_XmlParts", parts);
        }

        [HttpPost]
        public ActionResult ConvertToDslPartial(HttpPostedFileBase file)
        {
            if (file == null || file.ContentLength == 0)
                return PartialView("_Message", "Dosya seçilmedi.");

            string dsl;
            using (var s = file.InputStream)
            {
                var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
                dsl = ext == ".pptx" ? OpenXmlToDslConverter.ConvertPptxToDsl(s)
                    : ext == ".xlsx" ? OpenXmlToDslConverter.ConvertXlsxToDsl(s)
                    : "Bu uzantı için DSL dönüştürücü tanımlı değil.";
            }
            ViewBag.FileName = file.FileName;
            return PartialView("_Dsl", dsl);
        }

        [HttpPost]
        public ActionResult DownloadZip(HttpPostedFileBase file, bool includeRels = true)
        {
            if (file == null || file.ContentLength == 0) return new HttpStatusCodeResult(400, "Dosya yok");
            var parts = OpenXmlInspector.ExtractXmlParts(file.InputStream, includeRels);
            var zip = OpenXmlInspector.ToZip(parts);
            var name = Path.GetFileNameWithoutExtension(file.FileName) + "_xml.zip";
            return File(zip, "application/zip", name);
        }
    }
}

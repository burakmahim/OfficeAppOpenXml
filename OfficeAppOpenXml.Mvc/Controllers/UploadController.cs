using OfficeAppOpenXmlLibrary.Models;
using OfficeAppOpenXmlLibrary.Services;
using System.Configuration;
using System.IO;
using System.Web;
using System.Web.Mvc;

namespace OfficeAppOpenXml.Mvc.Controllers
{
    public class UploadController : Controller
    {
        private readonly FileService _svc;

        public UploadController()
        {
            string cs = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            _svc = new FileService(cs);
        }

        [HttpGet] public ActionResult Index() => View(_svc.List());

        [HttpGet] public ActionResult Upload() => View();

        [HttpPost, ValidateAntiForgeryToken]
        public ActionResult Upload(HttpPostedFileBase file)
        {
            if (file == null || file.ContentLength == 0)
            { ModelState.AddModelError("", "Dosya seçin."); return View(); }

            string fileName = Path.GetFileName(file.FileName);
            string mime = file.ContentType;
            using (MemoryStream ms = new MemoryStream()) { file.InputStream.CopyTo(ms); _svc.Upload(fileName, mime, ms.ToArray()); } 

            TempData["ok"] = "Dosya yüklendi.";
            return RedirectToAction("Index");
        }

        [HttpGet]
        public ActionResult Search(string q)
        {
            ViewBag.Query = q;
            System.Collections.Generic.IEnumerable<SearchResult> results = _svc.Search(q);
            return View(results);
        }

        [HttpGet]
        public ActionResult Download(int id)
        {
            (string FileName, string Mime, byte[] Bytes)? hit = _svc.GetFile(id);
            if (hit == null) return HttpNotFound();
            return File(hit.Value.Bytes, hit.Value.Mime, hit.Value.FileName);
        }

        [HttpPost]
        public ActionResult BackfillContentText()
        {
            int count = _svc.BackfillContentText();
            return Content($"Backfilled rows: {count}");
        }
    }
}

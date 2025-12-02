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
            var cs = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            _svc = new FileService(cs);
        }

        [HttpGet] public ActionResult Index() => View(_svc.List());

        [HttpGet] public ActionResult Upload() => View();

        [HttpPost, ValidateAntiForgeryToken]
        public ActionResult Upload(HttpPostedFileBase file)
        {
            if (file == null || file.ContentLength == 0)
            { ModelState.AddModelError("", "Dosya seçin."); return View(); }

            var fileName = Path.GetFileName(file.FileName);
            var mime = file.ContentType;
            using (var ms = new MemoryStream()) { file.InputStream.CopyTo(ms); _svc.Upload(fileName, mime, ms.ToArray()); }

            TempData["ok"] = "Dosya yüklendi.";
            return RedirectToAction("Index");
        }

        [HttpGet]
        public ActionResult Search(string q)
        {
            ViewBag.Query = q;
            var results = _svc.Search(q);
            return View(results);
        }

        [HttpGet]
        public ActionResult Download(int id)
        {
            var hit = _svc.GetFile(id);
            if (hit == null) return HttpNotFound();
            return File(hit.Value.Bytes, hit.Value.Mime, hit.Value.FileName);
        }

        [HttpPost]
        public ActionResult BackfillContentText()
        {
            var count = _svc.BackfillContentText();
            return Content($"Backfilled rows: {count}");
        }
    }
}

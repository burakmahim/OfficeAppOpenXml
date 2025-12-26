using Microsoft.AspNetCore.Mvc;
using OfficeAppOpenXmlLibrary.Services;

namespace OfficeAppOpenXml.Core.Controllers
{
    public class UploadController : Controller
    {
        private readonly FileService _svc;

        public UploadController(FileService svc)
        {
            _svc = svc;
        }

        // ---------------------------------------------------------
        // GET: /Upload  → Liste
        // ---------------------------------------------------------
        [HttpGet]
        public IActionResult Index()
        {
            var list = _svc.List();
            return View(list);
        }

        // ---------------------------------------------------------
        // GET: /Upload/Upload → Form
        // ---------------------------------------------------------
        [HttpGet]
        public IActionResult Upload()
        {
            return View();
        }

        // ---------------------------------------------------------
        // POST: /Upload/Upload → Yükleme
        // ---------------------------------------------------------
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Upload(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                ModelState.AddModelError("", "Dosya seçin.");
                return View();
            }

            using var ms = new MemoryStream();
            await file.CopyToAsync(ms);

            _svc.Upload(
                fileName: file.FileName,
                mime: file.ContentType,
                bytes: ms.ToArray()
            );

            TempData["ok"] = "Dosya yüklendi.";
            return RedirectToAction("Index");
        }

        // ---------------------------------------------------------
        // GET: /Upload/Search?q=...
        // ---------------------------------------------------------
        [HttpGet]
        public IActionResult Search(string q)
        {
            ViewBag.Query = q;
            var results = _svc.Search(q);
            return View(results);
        }

        // ---------------------------------------------------------
        // GET: /Upload/Download/5
        // ---------------------------------------------------------
        [HttpGet]
        public IActionResult Download(int id)
        {
            var hit = _svc.GetFile(id);
            if (hit == null) return NotFound();

            return File(hit.Value.Bytes, hit.Value.Mime, hit.Value.FileName);
        }

        // ---------------------------------------------------------
        // POST: /Upload/BackfillContentText
        // ---------------------------------------------------------
        [HttpPost]
        public IActionResult BackfillContentText()
        {
            var count = _svc.BackfillContentText();
            return Content($"Backfilled rows: {count}");
        }
    }
}

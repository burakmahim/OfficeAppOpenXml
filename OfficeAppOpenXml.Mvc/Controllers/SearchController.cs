using System.Collections.Generic;
using System.Web.Mvc;
using OfficeAppOpenXmlLibrary;

namespace OfficeAppOpenXml.Mvc.Controllers
{
    public class SearchController : Controller
    {
        private string folderPath = @"C:\Users\excalibur\Desktop\uploads";
        private string indexPath  = @"C:\Users\excalibur\Desktop\LuceneIndex";

        public ActionResult Index()
        {
            return View();
        }

		[HttpPost]
		public ActionResult Index(string Query)
		{
			List<SearchResult> results = new List<SearchResult>();

			if (!string.IsNullOrWhiteSpace(Query))
			{
				try
				{
					if (System.IO.Directory.Exists(folderPath))
					{
						LuceneSearchService searcher = new LuceneSearchService(indexPath);
						searcher.CreateIndex(folderPath);
						results = searcher.Search(Query);
					}
				}
				catch (System.Exception)
				{
				}
			}

			ViewBag.Query   = Query;
			ViewBag.Results = results;

			return View();
		}

	}
}

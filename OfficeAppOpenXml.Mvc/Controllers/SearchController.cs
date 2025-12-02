using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using OfficeAppOpenXmlLibrary;

namespace OfficeAppOpenXml.Mvc.Controllers
{
    public class SearchController : Controller
    {
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
					WindowsSearchService searcher = new WindowsSearchService();
					results = searcher.Search(Query);
				}
				catch (System.Exception ex)
				{
					ViewBag.Error = "Hata: " + ex.Message;
				}
			}

			ViewBag.Query   = Query;
			ViewBag.Results = results;

			return View();
		}

	}
}

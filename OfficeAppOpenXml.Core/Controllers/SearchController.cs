using Microsoft.AspNetCore.Mvc;
using OfficeAppOpenXmlLibrary;

namespace OfficeAppOpenXml.Core.Controllers
{
	public class SearchController : Controller
	{
		public IActionResult Index()
		{
			return View();
		}

		[HttpPost]
		public IActionResult Index(string Query)
		{
			List<SearchResult> results = new List<SearchResult>();

			if (!string.IsNullOrWhiteSpace(Query))
			{
				try
				{
					WindowsSearchService searcher = new WindowsSearchService();
					results = searcher.Search(Query);
				}
				catch (Exception ex)
				{
					ViewBag.Error = "Hata: " + ex.Message;
				}
			}

			ViewBag.Query = Query;
			ViewBag.Results = results;

			return View();
		}
	}
}

using Newtonsoft.Json;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Mvc;
using OfficeAppOpenXmlLibrary.Models;

namespace OfficeAppOpenXml.Mvc.Controllers
{
    public class SearchController : Controller
    {
        private readonly HttpClient _http = new HttpClient();

        public async Task<ActionResult> Index(string q)
        {
            if (string.IsNullOrWhiteSpace(q))
                return View(model: null);

            var apiUrl = "https://localhost:7280/api/search?q=" + q;
            var json = await _http.GetStringAsync(apiUrl);

            var results = JsonConvert.DeserializeObject<List<SearchResult>>(json);

            return View(results);
        }
    }
}

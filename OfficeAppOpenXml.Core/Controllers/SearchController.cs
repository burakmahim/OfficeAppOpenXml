using Microsoft.AspNetCore.Mvc;
using OfficeAppOpenXmlLibrary.Services;

namespace OfficeAppOpenXml.Core.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class SearchController : ControllerBase
    {
        private readonly FileService _fileService;

        public SearchController(FileService fileService)
        {
            _fileService = fileService;
        }

        [HttpGet]
        public IActionResult Get([FromQuery] string q)
        {
            if (string.IsNullOrWhiteSpace(q))
                return BadRequest("Arama terimi boş olamaz.");

            var results = _fileService.Search(
                query: q,
                top: 50,
                freeText: false,
                modelsId: null,
                ext: null,
                from: null,
                to: null
            );

            return Ok(results);
        }
    }
}

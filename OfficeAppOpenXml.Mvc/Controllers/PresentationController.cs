using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Net.NetworkInformation;
using System.Web;
using System.Web.Mvc;
using OfficeAppOpenXmlLibrary;
using OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents;
using OfficeAppOpenXmlLibrary.PowerPointOpenXmlComponents;

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
		public ActionResult UploadToDB(string xmlContent, string fileType)
		{
			try
			{
				byte[] fileBytes;
				string fileName;
				string mimeType;

				string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;


				using (SqlConnection conn = new SqlConnection(connectionString))
				{
					conn.Open();
					string sql = "INSERT INTO Files (FileName, FileType, MIMEType, FileContent, CreatedAt, UpdatedAt) VALUES (@name, @type, @mime, @fileBytes, @created, @updated)";
					using (SqlCommand cmd = new SqlCommand(sql, conn))
					{
						switch (fileType)
						{
							case "Excel":
								fileBytes = ExcelLibrary.CreateExcel(xmlContent);
								fileName  = "Rapor.xlsx";
								mimeType  = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
								break;

							case "PowerPoint":
								fileBytes = PowerPointLibrary.CreatePowerPointPresentation(xmlContent);
								fileName  = "Sunum.pptx";
								mimeType  = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
								break;

							case "Word":
								fileBytes = null; 
								fileName  = "Dokuman.docx";
								mimeType  = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
								break;

							default:
								return Content("Geçersiz dosya tipi");
						}

						cmd.Parameters.AddWithValue("@name"     , fileName );
						cmd.Parameters.AddWithValue("@type"     , fileType );
						cmd.Parameters.AddWithValue("@mime"     , mimeType );
						cmd.Parameters.AddWithValue("@fileBytes", fileBytes);
						cmd.Parameters.AddWithValue("@created"  , DateTime.Now);
						cmd.Parameters.AddWithValue("@updated"  , DateTime.Now);
						cmd.ExecuteNonQuery();
					}
				}

				return Content("Dosya başarıyla kaydedildi!");
			}
			catch (Exception ex)
			{
				return Content("Hata oluştu: " + ex.Message);
			}
		}

		[HttpPost]
		[ValidateInput(false)]
		public ActionResult DownloadPptx(string xmlContent)
		{
			try
			{
				byte[] pptBytes = PowerPointLibrary.CreatePowerPointPresentation(xmlContent);
				return File(pptBytes,
					"application/vnd.openxmlformats-officedocument.presentationml.presentation",
					"Sunum.pptx");
			}
			catch (Exception ex)
			{
				return Content("Hata oluştu: " + ex.Message);
			}
		}

		[HttpPost]
		[ValidateInput(false)]
		public ActionResult GenerateExcelFromXml(string xmlContent)
		{

			try
			{
				byte[] result = ExcelLibrary.CreateExcel(xmlContent);
				return File(result,
					"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
					"veriler.xlsx");
			}
			catch (Exception ex)
			{
				Response.StatusCode = 500;
				return Content("Excel oluşturulamadı: " + ex.Message);
			}
		}

	}
}
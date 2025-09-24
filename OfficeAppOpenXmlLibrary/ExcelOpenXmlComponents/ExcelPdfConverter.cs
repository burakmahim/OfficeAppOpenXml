#if NET48
using Syncfusion.ExcelToPdfConverter;
using Syncfusion.OfficeChartToImageConverter;
using Syncfusion.Pdf;
using Syncfusion.XlsIO;
#elif NET9_0
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Pdf;
#endif

namespace OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents
{
    public class ExcelPdfConverter
    {
		public static byte[] ConvertToPdf(string xmlContent)
		{
			byte[] excelBytes = ExcelLibrary.CreateExcel(xmlContent);

			using MemoryStream ms = new MemoryStream(excelBytes);

#if NET48
            using ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;

            IWorkbook workbook = application.Workbooks.Open(ms);

            ExcelToPdfConverter converter = new ExcelToPdfConverter(workbook);

            ExcelToPdfConverterSettings settings = new ExcelToPdfConverterSettings
            {
                LayoutOptions = LayoutOptions.FitSheetOnOnePage
            };
            converter.ChartToImageConverter = new ChartToImageConverter() as IChartToImageConverter;

            PdfDocument pdfDocument = converter.Convert(settings);

            using MemoryStream outMs = new MemoryStream();
            pdfDocument.Save(outMs);
            return outMs.ToArray();

#elif NET9_0
			using ExcelEngine excelEngine = new ExcelEngine();
			IApplication application = excelEngine.Excel;
			application.DefaultVersion = ExcelVersion.Xlsx;

			IWorkbook workbook = application.Workbooks.Open(ms);

			XlsIORendererSettings settings = new XlsIORendererSettings
			{
				LayoutOptions = LayoutOptions.FitSheetOnOnePage
			};

			XlsIORenderer renderer = new XlsIORenderer();
			PdfDocument pdfDocument = renderer.ConvertToPDF(workbook, settings);

			using MemoryStream pdfStream = new MemoryStream();
			pdfDocument.Save(pdfStream);

			return pdfStream.ToArray();

			//#else
			//    throw new PlatformNotSupportedException("Bu platform desteklenmiyor.");
#endif
		}
	}
}

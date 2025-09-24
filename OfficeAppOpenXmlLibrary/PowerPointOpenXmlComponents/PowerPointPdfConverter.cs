#if NET48
using Syncfusion.Presentation;
using Syncfusion.PresentationToPdfConverter;
using Syncfusion.Pdf;
#elif NET9_0
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Syncfusion.Pdf;
#endif

namespace OfficeAppOpenXmlLibrary.PowerPointOpenXmlComponents
{
    public class PowerPointPdfConverter
    {
		public static byte[] ConvertToPdf(string xmlContent)
		{
			byte[] pptxBytes = PowerPointLibrary.CreatePowerPointPresentation(xmlContent);
			using MemoryStream ms = new MemoryStream(pptxBytes);

#if NET48
            using (IPresentation presentation = Presentation.Open(ms))
            {
                Syncfusion.OfficeChartToImageConverter.ChartToImageConverter chartToImageConverter = new Syncfusion.OfficeChartToImageConverter.ChartToImageConverter();

                presentation.ChartToImageConverter = chartToImageConverter;

                PresentationToPdfConverterSettings settings = new PresentationToPdfConverterSettings();
                settings.ShowHiddenSlides = false;

                using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(presentation, settings))
                {
                    using MemoryStream outMs = new MemoryStream();
                    pdfDocument.Save(outMs);
                    return outMs.ToArray();
                }
            }
#elif NET9_0
			using (IPresentation presentation = Presentation.Open(ms))
			{
				PdfDocument pdfDocument = PresentationToPdfConverter.Convert(presentation);

				using MemoryStream outMs = new MemoryStream();
				pdfDocument.Save(outMs);
				return outMs.ToArray();
			}
#else
           throw new PlatformNotSupportedException("Bu platform desteklenmiyor.");
#endif
		}

	}
}

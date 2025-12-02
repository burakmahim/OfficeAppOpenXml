using System.Text;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace OfficeAppOpenXmlLibrary.Models
{
    public class PdfExtractor
    {
        public string ExtractPdfText(byte[] bytes)
        {
            using var pdf = PdfDocument.Open(bytes);

            var sb = new StringBuilder();

            foreach (var page in pdf.GetPages())
            {
                sb.AppendLine(page.Text);
            }

            return sb.ToString();
        }
    }
}

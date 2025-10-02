using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAppOpenXmlLibrary.OpenXmlFileUpload
{
    public static class OpenXmlToDslConverter
    {
        public static string ConvertPptxToDsl(Stream pptx)
        {
            using PresentationDocument doc = PresentationDocument.Open(pptx, false);
            PresentationPart presPart = doc.PresentationPart!;
            XElement root = new XElement("presentation");

            foreach (var sid in presPart.Presentation.SlideIdList.Elements<SlideId>())
            {
                SlidePart sp = (SlidePart)presPart.GetPartById(sid.RelationshipId);
                XElement slideEl = new XElement("slide", new XAttribute("layout", "default"), new XAttribute("addFooter", "false"));

                foreach (Shape shape in sp.Slide.Descendants<Shape>())
                {
                    string text = string.Join(" ", shape.TextBody?.Descendants<A.Text>().Select(t => t.Text) ?? []);
                    PlaceholderShape? ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
                    if (string.IsNullOrWhiteSpace(text)) continue;
                    slideEl.Add(ph?.Type?.Value == PlaceholderValues.Title ? new XElement("title", text)
                                                                          : new XElement("content", text));
                }

                foreach (DocumentFormat.OpenXml.Presentation.Picture pic in sp.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>())
                {
                    string? relId = pic.BlipFill?.Blip?.Embed?.Value;
                    if (relId == null) continue;
                    ImagePart imgPart = (ImagePart)sp.GetPartById(relId);
                    using Stream ims = imgPart.GetStream();
                    using MemoryStream ms = new MemoryStream();
                    ims.CopyTo(ms);
                    string b64 = Convert.ToBase64String(ms.ToArray());
                    slideEl.Add(new XElement("image",
                        new XAttribute("path", $"data:{imgPart.ContentType};base64,{b64}"),
                        new XAttribute("x", "1"), new XAttribute("y", "1"),
                        new XAttribute("width", "4"), new XAttribute("height", "3")));
                }

                root.Add(slideEl);
            }

            return new XDocument(root).ToString(SaveOptions.DisableFormatting);
        }

        public static string ConvertXlsxToDsl(Stream xlsx)
        {
            using SpreadsheetDocument doc = SpreadsheetDocument.Open(xlsx, false);
            WorkbookPart wbPart = doc.WorkbookPart!;
            SharedStringTable? sst = wbPart.SharedStringTablePart?.SharedStringTable;
            XElement root = new XElement("workbook");

            foreach (Sheet s in wbPart.Workbook.Sheets!.Elements<Sheet>())
            {
                WorksheetPart wsp = (WorksheetPart)wbPart.GetPartById(s.Id!);
                XElement wsEl = new XElement("worksheet", new XAttribute("name", s.Name!));
                XElement tableEl = new XElement("table", new XAttribute("name", $"{s.Name}-Data"));

                foreach (Row row in wsp.Worksheet.Descendants<Row>())
                {
                    XElement rowEl = new XElement("row");
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        string val = cell.CellValue?.InnerText ?? "";
                        if (cell.DataType?.Value == CellValues.SharedString && sst != null && int.TryParse(val, out int i) && i >= 0 && i < sst.Count())
                            val = sst.ElementAt(i).InnerText;
                        rowEl.Add(new XElement("cell", val));
                    }
                    tableEl.Add(rowEl);
                }

                wsEl.Add(tableEl);
                root.Add(wsEl);
            }

            return new XDocument(root).ToString(SaveOptions.DisableFormatting);
        }
    }
}

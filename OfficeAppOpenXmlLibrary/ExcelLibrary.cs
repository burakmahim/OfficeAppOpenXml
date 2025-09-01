using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents;

namespace OfficeAppOpenXmlLibrary
{
    public class ExcelLibrary
    {
        public static byte[] CreateExcel(string xmlContent)
        {
            if (string.IsNullOrWhiteSpace(xmlContent))
                throw new ArgumentException("xmlContent boş.", nameof(xmlContent));

            using var memoryStream = new MemoryStream();
            using (var document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var sheets = workbookPart.Workbook.GetFirstChild<Sheets>() ?? workbookPart.Workbook.AppendChild(new Sheets());

                var doc = XDocument.Parse(xmlContent);
                int sheetCounter = 1;

                // İlk iş: önceki kayıtları temizle
                TableComponent.TableRegistry.Clear();

                foreach (var sheetElement in doc.Root.Elements("sheet"))
                {
                    string sheetName = sheetElement.Attribute("name")?.Value ?? $"Sheet{sheetCounter}";

                    // WorksheetPart + Worksheet + SheetData
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var worksheet = new Worksheet();
                    worksheet.Append(new SheetData());

                    // Tablolar
                    foreach (var tableElement in sheetElement.Elements("table"))
                    {
                        var placement = TableComponent.AddTable(tableElement, worksheet, sheetName);
                        // placement.RangeA1 ve placement.StartCell hazır
                    }

                    // Worksheet'i ata
                    worksheetPart.Worksheet = worksheet;

                    // Grafikleri işle (ÖRNEK: sadece tablo referansı varsa)
                    foreach (var chartEl in sheetElement.Elements("chart"))
                    {
                        // Eğer tablo referansı verilmişse:
                        var tableRef = (string?)chartEl.Attribute("tableRef");
                        if (!string.IsNullOrWhiteSpace(tableRef) && TableComponent.TableRegistry.TryGetValue(tableRef, out var pl))
                        {
                            string catCol = (string?)chartEl.Attribute("categoryColumn") ?? "A";
                            string valCols = (string?)chartEl.Attribute("valueColumns") ?? "B";

                            var chartDef = ChartXmlHelpers.BuildChartDefinitionFromTable(
                                chartEl,
                                pl,
                                sheetName,      // <--- worksheetPart değil, sayfa adı
                                catCol,
                                valCols);

                            ExcelChartSupport.AddChart(
                                worksheetPart,
                                chartDef,
                                topLeftCell: (string?)chartEl.Attribute("topLeftCell") ?? "B2",
                                bottomRightCell: (string?)chartEl.Attribute("bottomRightCell") ?? "M20");
                        }
                        else
                        {
                            var chartDef = ChartXmlHelpers.BuildChartDefinitionInline(chartEl);
                            ExcelChartSupport.AddChart(
                                worksheetPart,
                                chartDef,
                                topLeftCell: (string?)chartEl.Attribute("topLeftCell") ?? "B2",
                                bottomRightCell: (string?)chartEl.Attribute("bottomRightCell") ?? "M20");
                        }

                    }

                    // Workbook/Sheets kaydı
                    var sheet = new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = (uint)sheetCounter,
                        Name = sheetName
                    };
                    sheets.Append(sheet);

                    sheetCounter++;
                }

                // calcPr
                var calcPr = workbookPart.Workbook.Elements<CalculationProperties>().FirstOrDefault();
                if (calcPr == null)
                {
                    calcPr = new CalculationProperties { FullCalculationOnLoad = true };
                    var sheetsEl = workbookPart.Workbook.GetFirstChild<Sheets>();
                    if (sheetsEl != null)
                        workbookPart.Workbook.InsertAfter(calcPr, sheetsEl);
                    else
                        workbookPart.Workbook.Append(calcPr);
                }
                else
                {
                    calcPr.FullCalculationOnLoad = true;
                }

#if DEBUG
                var validator = new OpenXmlValidator(FileFormatVersions.Office2013);
                var errors = validator.Validate(document).Take(5).ToList();
                if (errors.Count > 0)
                {
                    var msg = string.Join(Environment.NewLine, errors.Select(e =>
                        $"Part: {(e.Part != null ? e.Part.Uri.ToString() : "(package)")} | Path: {e.Path} | Desc: {e.Description}"
                    ));
                    throw new InvalidOperationException("OpenXML Validation Errors:\n" + msg);
                }
#endif
            }

            return memoryStream.ToArray();
        }
    }
}

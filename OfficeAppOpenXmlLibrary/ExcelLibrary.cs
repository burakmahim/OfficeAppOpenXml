using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents; // TableComponent, ChartSpec, ExcelChartHelper

namespace OfficeAppOpenXmlLibrary
{
    public class ExcelLibrary
    {
        /// <summary>
        /// Verilen XML'e göre XLSX üretir.
        /// Şema: <excel><sheet name="..."><table startCell="A1">...</table><chart ... /></sheet>...</excel>
        /// </summary>
        public static byte[] CreateExcel(string xmlContent)
        {
            if (string.IsNullOrWhiteSpace(xmlContent))
                throw new ArgumentException("xmlContent boş.", nameof(xmlContent));

            using (var memoryStream = new MemoryStream())
            {
                using (var document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
                {
                    // --- Workbook ---
                    var workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    // Sheets: varsa kullan, yoksa tek bir örnek oluştur (aynı örneği iki kez eklememek için)
                    var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
                    if (sheets == null)
                        sheets = workbookPart.Workbook.AppendChild(new Sheets());

                    // --- XML yükle ---
                    var doc = XDocument.Parse(xmlContent);
                    int sheetCounter = 1;

                    foreach (var sheetElement in doc.Root.Elements("sheet"))
                    {
                        string sheetName = sheetElement.Attribute("name")?.Value ?? $"Sheet{sheetCounter}";

                        // 1) WorksheetPart + Worksheet + SheetData
                        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        var worksheet = new Worksheet();
                        var sheetData = new SheetData();
                        worksheet.Append(sheetData);

                        // 2) Tablolar
                        foreach (var tableElement in sheetElement.Elements("table"))
                        {
                            TableComponent.AddTable(tableElement, worksheet, out _, out _);
                        }

                        // 3) Worksheet'i PART'a ATA (grafiklerden ÖNCE!)
                        worksheetPart.Worksheet = worksheet;

                        // 4) Grafikleri ekle (varsa)
                        foreach (var chartEl in sheetElement.Elements("chart"))
                        {
                            try
                            {
                                var type = (string?)chartEl.Attribute("type") ?? "column";
                                var title = (string?)chartEl.Attribute("title") ?? null;
                                var from = (string?)chartEl.Attribute("from") ?? "D2";
                                var to = (string?)chartEl.Attribute("to") ?? "L20";

                                ChartSpec spec;

                                // --- INLINE VERİ VAR MI? (series/point) ---
                                var seriesEls = chartEl.Elements("series");
                                if (seriesEls.Any())
                                {
                                    var categories = new List<string>();
                                    var valuesPerSeries = new List<List<double?>>();
                                    var names = new List<string>();

                                    bool first = true;
                                    foreach (var serEl in seriesEls)
                                    {
                                        var name = (string?)serEl.Attribute("name") ?? $"Series {names.Count + 1}";
                                        names.Add(name);

                                        // İlk seride kategorileri sırayla topla
                                        if (first)
                                        {
                                            categories.Clear();
                                            foreach (var pt in serEl.Elements("point"))
                                                categories.Add((string?)pt.Attribute("category") ?? "");
                                            first = false;
                                        }

                                        // Bu seri için değerleri kategorilere göre sırala
                                        var vals = new List<double?>(new double?[categories.Count]);
                                        // varsayılan 0 yazmak istersen: Enumerable.Repeat<double?>(0, categories.Count).ToList()

                                        foreach (var pt in serEl.Elements("point"))
                                        {
                                            var cat = (string?)pt.Attribute("category") ?? "";
                                            var valStr = (string?)pt.Attribute("value") ?? "0";
                                            if (!double.TryParse(valStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var dv))
                                                dv = 0;

                                            int idx = categories.IndexOf(cat);
                                            if (idx >= 0) vals[idx] = dv;
                                        }

                                        // Boş kalanlar 0 olsun:
                                        for (int i = 0; i < vals.Count; i++)
                                            vals[i] ??= 0;

                                        valuesPerSeries.Add(vals);
                                    }

                                    spec = new ChartSpec
                                    {
                                        Type = type,
                                        Title = title,
                                        FromCell = from,
                                        ToCell = to,
                                        CategoriesInline = categories,
                                        ValuesInline = valuesPerSeries,
                                        SeriesNames = names,
                                        ShowLegend = true
                                    };
                                }
                                else
                                {
                                    // ESKİ A1 ŞEMASI
                                    var categories = (string?)chartEl.Attribute("categories") ?? "'Sheet1'!$A$2:$A$6";
                                    var valuesAttr = (string?)chartEl.Attribute("values") ?? "'Sheet1'!$B$2:$B$6";
                                    var valueRanges = valuesAttr.Split(';').Select(s => s.Trim()).Where(s => s.Length > 0).ToList();
                                    var seriesAttr = (string?)chartEl.Attribute("series") ?? "";
                                    var seriesNames = seriesAttr.Split(';').Select(s => s.Trim()).Where(s => s.Length > 0).ToList();
                                    if (seriesNames.Count == 0) seriesNames = null;

                                    spec = new ChartSpec
                                    {
                                        Type = type,
                                        Title = title,
                                        FromCell = from,
                                        ToCell = to,
                                        CategoryRange = categories,
                                        ValueRanges = valueRanges,
                                        SeriesNames = seriesNames
                                    };
                                }

                                ExcelChartHelper.AddChart(worksheetPart, spec, spec.FromCell!, spec.ToCell!);
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine("Chart error: " + ex);
                            }
                        }

                        // 5) Workbook'a sheet kaydı
                        var sheet = new Sheet
                        {
                            Id = workbookPart.GetIdOfPart(worksheetPart),
                            SheetId = (uint)sheetCounter,
                            Name = sheetName
                        };
                        sheets.Append(sheet);

                        sheetCounter++;
                    }

                    // 6) calcPr: tek kopya ve Sheets'ten sonra
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

                    workbookPart.Workbook.Save();

#if DEBUG
                    // Geliştirme sırasında paket doğrulama (en fazla 5 hata göster)
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

        /// <summary>
        /// PDF'e çevirmeden önce formülleri "cached" değere dönüştürmek istersen kullan.
        /// </summary>
        public static byte[] CreateExcelForPdf(string xmlContent)
        {
            var bytes = CreateExcel(xmlContent);

            using (var ms = new MemoryStream(bytes))
            using (var doc = SpreadsheetDocument.Open(ms, true))
            {
                foreach (var wsPart in doc.WorkbookPart.WorksheetParts)
                    FlattenFormulas(wsPart.Worksheet);

                doc.WorkbookPart.Workbook.Save();
                return ms.ToArray();
            }
        }

        /// <summary>
        /// Hücrelerde formül varsa ve CellValue (cached) mevcutsa, formülü kaldırıp değeri bırakır.
        /// </summary>
        private static void FlattenFormulas(Worksheet ws)
        {
            var sd = ws.GetFirstChild<SheetData>();
            if (sd == null) return;

            foreach (var row in sd.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    if (cell.CellFormula == null) continue;

                    var v = cell.CellValue?.Text;
                    if (string.IsNullOrWhiteSpace(v)) continue; // cached yoksa dokunma

                    // Formülü kaldır, değeri bırak
                    cell.CellFormula = null;

                    double _;
                    if (double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                    {
                        // Sayı
                        cell.DataType = null;
                        cell.InlineString = null;
                        cell.CellValue = new CellValue(v);
                    }
                    else
                    {
                        // Metin
                        cell.CellValue = null;
                        cell.DataType = CellValues.InlineString;
                        cell.InlineString = new InlineString(new Text(v) { Space = SpaceProcessingModeValues.Preserve });
                    }
                }
            }
        }
    }
}

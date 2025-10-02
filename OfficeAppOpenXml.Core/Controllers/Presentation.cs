using Microsoft.AspNetCore.Mvc;
using OfficeAppOpenXmlLibrary;
using OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents;

namespace OfficeAppOpenXml.Core.Controllers
{
    public class PresentationController : Controller
    {
        [HttpGet]
        public  IActionResult   Index()                                                 
        {
            ViewBag.XmlContent = "";
            return View();
        }

        [HttpPost]
        public  IActionResult   DownloadPptx            ([FromForm] string xmlContent)  
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
        public  IActionResult   GenerateExcelFromXml    ([FromForm] string xmlContent)  
        {
            if (string.IsNullOrWhiteSpace(xmlContent))
                return BadRequest("XML boş olamaz.");

            try
            {
                byte[] result = ExcelLibrary.CreateExcel(xmlContent);
                return File(result,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "veriler.xlsx");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Excel oluşturulamadı: {ex.Message}");
            }
        }

        [HttpPost]
        public  IActionResult   UploadExcelAndViewXml   (IFormFile excelFile)           
        {

                if (excelFile == null || excelFile.Length == 0)
                {
                    return Json(new { success = false, message = "Dosya seçilmedi" });
                }

                using (MemoryStream memoryStream = new MemoryStream())
                {
                    excelFile.CopyTo(memoryStream);
                    byte[] fileBytes = memoryStream.ToArray();

                    string xmlContent = ConvertExcelToCustomXml(fileBytes);

                    return Json(new { success = true, xmlContent = xmlContent });
                }
        }

        private string          ConvertExcelToCustomXml (byte[] excelData)              
        {
            using (MemoryStream memoryStream = new MemoryStream(excelData))
            {
                using (DocumentFormat.OpenXml.Packaging.SpreadsheetDocument document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(memoryStream, false))
                {
                    DocumentFormat.OpenXml.Packaging.WorkbookPart workbookPart = document.WorkbookPart;
                    if (workbookPart?.Workbook?.Sheets == null)
                        return "<workbook><error>Excel okunamadı</error></workbook>";

                    System.Text.StringBuilder xmlBuilder = new System.Text.StringBuilder();
                    xmlBuilder.AppendLine("<workbook>");

                    foreach (DocumentFormat.OpenXml.Spreadsheet.Sheet sheet in workbookPart.Workbook.Sheets)
                    {
                        DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart    = (DocumentFormat.OpenXml.Packaging.WorksheetPart)workbookPart.GetPartById(sheet.Id!);
                        DocumentFormat.OpenXml.Spreadsheet.Worksheet   worksheet        = worksheetPart.Worksheet;

                        xmlBuilder.AppendLine($"<worksheet name=\"{sheet.Name?.Value ?? "Sheet"}\">");

                        if (worksheetPart.TableDefinitionParts.Any())
                        {
							ConvertWorksheetToTable(worksheet, workbookPart, xmlBuilder);
						}

						if (worksheetPart.DrawingsPart?.ChartParts.Any() == true)
                        {
                            ConvertChartsToXml(worksheetPart, workbookPart, xmlBuilder);
                        }

                        xmlBuilder.AppendLine("</worksheet>");
                    }

                    xmlBuilder.AppendLine("</workbook>");
                    return xmlBuilder.ToString();
                }
            }
        }

        private void            ConvertWorksheetToTable (DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet, DocumentFormat.OpenXml.Packaging.WorkbookPart workbookPart, System.Text.StringBuilder xmlBuilder)
        {
            DocumentFormat.OpenXml.Spreadsheet.SheetData sheetData = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
            if (sheetData == null) return;

            xmlBuilder.AppendLine("<table>");

            System.Collections.Generic.IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Row> rows = sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().OrderBy(r => r.RowIndex?.Value ?? 0);

            foreach (DocumentFormat.OpenXml.Spreadsheet.Row row in rows)
            {
                xmlBuilder.AppendLine("<row>");

                System.Collections.Generic.IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Cell> cells = row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().OrderBy(c => GetColumnIndex(c.CellReference?.Value ?? "A1"));

                foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in cells)
                {
                    string cellValue   = GetCellValue(cell, workbookPart);
                    string cellFormula = cell.CellFormula?.Text ?? "";

                    if (!string.IsNullOrEmpty(cellFormula))
                    {
                        xmlBuilder.AppendLine($"<cell formula=\"{System.Security.SecurityElement.Escape(cellFormula)}\">{System.Security.SecurityElement.Escape(cellValue)}</cell>");
                    }
                    else
                    {
                        xmlBuilder.AppendLine($"<cell>{System.Security.SecurityElement.Escape(cellValue)}</cell>");
                    }
                }

                xmlBuilder.AppendLine("</row>");
            }

            xmlBuilder.AppendLine("</table>");
        }

        private void            ConvertChartsToXml      (DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart, DocumentFormat.OpenXml.Packaging.WorkbookPart workbookPart, System.Text.StringBuilder xmlBuilder)
        {
            foreach (DocumentFormat.OpenXml.Packaging.ChartPart chartPart in worksheetPart.DrawingsPart.ChartParts)
            {
                DocumentFormat.OpenXml.Drawing.Charts.ChartSpace chartSpace = chartPart.ChartSpace;
                DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartSpace.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();

                string chartTitle = chart?.Title?.ChartText?.RichText?.InnerText ?? "";
                string chartType = GetChartType(chart);

                xmlBuilder.AppendLine($"<chart type=\"{chartType}\" title=\"{System.Security.SecurityElement.Escape(chartTitle)}\">");

                // Category Axis 
                if (chart?.PlotArea?.Elements<DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis>().Any() == true)
                {
                    DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis categoryAxis = chart.PlotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis>().FirstOrDefault();
                    
                    string categoryAxisTitle = categoryAxis?.Title?.ChartText?.RichText?.InnerText ?? "";

                    xmlBuilder.AppendLine ("    <category-axis>");
                    AddFormatStructure    (xmlBuilder);
					AddTextFormatStructure(xmlBuilder);
					AddTitleStructure     (xmlBuilder, categoryAxisTitle);
                    xmlBuilder.AppendLine ("    </category-axis>");
                }

                // Value Axis
                if (chart?.PlotArea?.Elements<DocumentFormat.OpenXml.Drawing.Charts.ValueAxis>().Any() == true)
                {
                    DocumentFormat.OpenXml.Drawing.Charts.ValueAxis valueAxis = chart.PlotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.ValueAxis>().FirstOrDefault();
                    
                    string valueAxisTitle = valueAxis?.Title?.ChartText?.RichText?.InnerText ?? "";

                    xmlBuilder.AppendLine ("    <value-axis>");
					AddFormatStructure    (xmlBuilder);
					AddTextFormatStructure(xmlBuilder);
					AddTitleStructure     (xmlBuilder, valueAxisTitle);
					xmlBuilder.AppendLine ("    </value-axis>");
                }

                // Legend  
                if (chart?.Legend != null)
                {
                    xmlBuilder.AppendLine ("    <legend>");
					AddFormatStructure    (xmlBuilder);
					AddTextFormatStructure(xmlBuilder);
					xmlBuilder.AppendLine ("    </legend>");
                }

                // Grid  
                if (chart?.PlotArea?.Elements<DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis>().Any(axis => axis.MajorGridlines != null || axis.MinorGridlines != null) == true ||
                    chart?.PlotArea?.Elements<DocumentFormat.OpenXml.Drawing.Charts.ValueAxis>   ().Any(axis => axis.MajorGridlines != null || axis.MinorGridlines != null) == true)
                {
                    xmlBuilder.AppendLine("    <grid>");
                    xmlBuilder.AppendLine("        <horizontal-format />");
                    xmlBuilder.AppendLine("        <vertical-format />");
                    xmlBuilder.AppendLine("    </grid>");
                }

                // Data Labels 
                bool hasDataLabels = false;
                if (chart?.PlotArea != null)
                {
                    hasDataLabels = chart.PlotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.BarChart> ().Any(bc => bc.Elements().Any(s => s.Elements().Any(e => e.LocalName == "dLbls"))) ||
                                    chart.PlotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.LineChart>().Any(lc => lc.Elements().Any(s => s.Elements().Any(e => e.LocalName == "dLbls"))) ||
                                    chart.PlotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.PieChart> ().Any(pc => pc.Elements().Any(s => s.Elements().Any(e => e.LocalName == "dLbls")));
                }

                if (hasDataLabels)
                {
                    xmlBuilder.AppendLine("    <value-labels />");
                }

                // Series  
                bool hasSeries = false;
                if (chart?.PlotArea != null)
                {
                    hasSeries = chart.PlotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.BarChart> ().Any() ||
                                chart.PlotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.LineChart>().Any() ||
                                chart.PlotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.PieChart> ().Any() ||
                                chart.PlotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.AreaChart>().Any();
                }

                if (hasSeries)
                {
                    xmlBuilder.AppendLine("    <series>");
                    xmlBuilder.AppendLine("        <points>");
                    xmlBuilder.AppendLine("            <point>");
                    xmlBuilder.AppendLine("                <format />");
                    xmlBuilder.AppendLine("                <value-labels />");
                    xmlBuilder.AppendLine("            </point>");
                    xmlBuilder.AppendLine("        </points>");
					AddFormatStructure(xmlBuilder);
					xmlBuilder.AppendLine("        <marker />");
                    xmlBuilder.AppendLine("        <title />");
                    xmlBuilder.AppendLine("        <value-labels />");
                    xmlBuilder.AppendLine("    </series>");
                }

                // Layout 
                if (chart?.PlotArea?.Layout != null)
                {
                    xmlBuilder.AppendLine("    <layout />");
                }

                // 3D View
                if (chart?.View3D != null)
                {
                    xmlBuilder.AppendLine("    <view-3d>");
					AddFormatStructure(xmlBuilder);
					xmlBuilder.AppendLine("        <floor-format />");
                    xmlBuilder.AppendLine("        <back-wall-format />");
                    xmlBuilder.AppendLine("        <side-wall-format />");
                    xmlBuilder.AppendLine("    </view-3d>");
                }

                // Plot Area 
                if (chart?.PlotArea != null)
                {
                    xmlBuilder.AppendLine("    <plot-area>");
					AddFormatStructure(xmlBuilder);
					xmlBuilder.AppendLine("    </plot-area>");
                }

				// Chart seviyesi formatlar
				AddFormatStructure    (xmlBuilder);
				AddTextFormatStructure(xmlBuilder);
				xmlBuilder.AppendLine ("    <marker />");
				AddTitleStructure     (xmlBuilder, chartTitle);
                xmlBuilder.AppendLine ("    <series-format />");

                // Data Table
                bool hasDataTable = false;
                if (chart?.PlotArea != null)
                {
                    hasDataTable = chart.PlotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.DataTable>().Any();
                }

                if (hasDataTable)
                {
                    xmlBuilder.AppendLine ("    <data-table>");
					AddFormatStructure    (xmlBuilder);
					AddTextFormatStructure(xmlBuilder);
					xmlBuilder.AppendLine ("    </data-table>");
                }

                xmlBuilder.AppendLine("</chart>");
            }
        }

        private void            AddFormatStructure      (System.Text.StringBuilder xmlBuilder)
        {
            xmlBuilder.AppendLine("<format>");
            xmlBuilder.AppendLine("    <fill>");
            xmlBuilder.AppendLine("        <solid />");
            xmlBuilder.AppendLine("        <pattern>");
            xmlBuilder.AppendLine("            <stop />");
            xmlBuilder.AppendLine("        </pattern>");
            xmlBuilder.AppendLine("        <gradient>");
            xmlBuilder.AppendLine("            <stop />");
            xmlBuilder.AppendLine("        </gradient>");
            xmlBuilder.AppendLine("    </fill>");
            xmlBuilder.AppendLine("    <line>");
            xmlBuilder.AppendLine("        <solid />");
            xmlBuilder.AppendLine("        <gradient>");
            xmlBuilder.AppendLine("            <stop />");
            xmlBuilder.AppendLine("        </gradient>");
            xmlBuilder.AppendLine("    </line>");
            xmlBuilder.AppendLine("    <effects>");
            xmlBuilder.AppendLine("        <shadow />");
            xmlBuilder.AppendLine("        <glow />");
            xmlBuilder.AppendLine("        <soft-edges />");
            xmlBuilder.AppendLine("        <reflection />");
            xmlBuilder.AppendLine("        <format-3d>");
            xmlBuilder.AppendLine("            <bevel />");
            xmlBuilder.AppendLine("            <lighting />");
            xmlBuilder.AppendLine("        </format-3d>");
            xmlBuilder.AppendLine("    </effects>");
            xmlBuilder.AppendLine("</format>");
        }

        private void            AddTextFormatStructure  (System.Text.StringBuilder xmlBuilder)
        {
            xmlBuilder.AppendLine("<text-format>");
            xmlBuilder.AppendLine("    <fill>");
            xmlBuilder.AppendLine("        <solid />");
            xmlBuilder.AppendLine("        <pattern>");
            xmlBuilder.AppendLine("            <stop />");
            xmlBuilder.AppendLine("        </pattern>");
            xmlBuilder.AppendLine("        <gradient>");
            xmlBuilder.AppendLine("            <stop />");
            xmlBuilder.AppendLine("        </gradient>");
            xmlBuilder.AppendLine("    </fill>");
            xmlBuilder.AppendLine("    <line>");
            xmlBuilder.AppendLine("        <solid />");
            xmlBuilder.AppendLine("        <gradient>");
            xmlBuilder.AppendLine("            <stop />");
            xmlBuilder.AppendLine("        </gradient>");
            xmlBuilder.AppendLine("    </line>");
            xmlBuilder.AppendLine("    <effects>");
            xmlBuilder.AppendLine("        <shadow />");
            xmlBuilder.AppendLine("        <glow />");
            xmlBuilder.AppendLine("        <soft-edges />");
            xmlBuilder.AppendLine("        <reflection />");
            xmlBuilder.AppendLine("        <format-3d>");
            xmlBuilder.AppendLine("            <bevel />");
            xmlBuilder.AppendLine("            <lighting />");
            xmlBuilder.AppendLine("        </format-3d>");
            xmlBuilder.AppendLine("    </effects>");
            xmlBuilder.AppendLine("    <font />");
            xmlBuilder.AppendLine("</text-format>");
        }

        private void            AddTitleStructure       (System.Text.StringBuilder xmlBuilder, string titleText = "")
        {
            xmlBuilder.AppendLine($"<title text=\"{System.Security.SecurityElement.Escape(titleText)}\">");
            AddTextFormatStructure(xmlBuilder);
            xmlBuilder.AppendLine("</title>");
        }

        private string          GetCellValue            (DocumentFormat.OpenXml.Spreadsheet.Cell cell,           DocumentFormat.OpenXml.Packaging.WorkbookPart workbookPart)                                      
        {
            if (cell.CellValue == null)
                return "";

            string value = cell.CellValue.Text;

            if (cell.DataType != null && cell.DataType.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString)
            {
				DocumentFormat.OpenXml.Spreadsheet.SharedStringTable? sharedStrings  = workbookPart.SharedStringTablePart?.SharedStringTable;
                if (sharedStrings != null && int.TryParse(value, out int index))
                {
                    return sharedStrings.Elements<DocumentFormat.OpenXml.Spreadsheet.SharedStringItem>().ElementAtOrDefault(index)?.InnerText ?? "";
                }
            }

            return value;
        }

        private int             GetColumnIndex          (string cellReference)          
        {
            if (string.IsNullOrEmpty(cellReference))
                return 0;

            var columnName = "";
            for (int i = 0; i < cellReference.Length; i++)
            {
                if (char.IsLetter(cellReference[i]))
                    columnName += cellReference[i];
                else
                    break;
            }

            int columnIndex = 0;
            for (int i = 0; i < columnName.Length; i++)
            {
                columnIndex = columnIndex * 26 + (columnName[i] - 'A' + 1);
            }
            return columnIndex - 1;
        }
    }
}

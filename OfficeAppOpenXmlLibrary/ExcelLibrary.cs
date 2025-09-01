using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml.Linq;
using System.IO;
using OfficeAppOpenXmlLibrary.PowerPointOpenXmlComponents;
using OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents;

namespace OfficeAppOpenXmlLibrary
{
    public class ExcelLibrary
    {
        public static byte[] CreateExcel(string xmlContent)
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    Sheets sheets = new Sheets();
                    workbookPart.Workbook.Append(sheets);

                    XDocument doc = XDocument.Parse(xmlContent);
                    int sheetCounter = 1;

                    foreach (XElement sheetElement in doc.Root.Elements("sheet"))
                    {
                        string sheetName = sheetElement.Attribute("name")?.Value ?? $"Sayfa {sheetCounter}";

                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        Worksheet worksheet = new Worksheet();
                        SheetData sheetData = new SheetData();

                        worksheet.Append(sheetData);

                        int rowCount = 0;
                        int colCount = 0;
                        int currentRow = 1;


                        foreach (XElement tableElement in sheetElement.Elements("table"))
                        {
                            TableComponent.AddTable(tableElement, worksheet, out rowCount, out colCount, currentRow);
                            currentRow += rowCount + 1;
                        }

                        worksheetPart.Worksheet = worksheet;

                        foreach (XElement chartElement in sheetElement.Elements("chart"))
                        {
                            ChartComponent.AddChart(worksheetPart, chartElement);
                        }

                        Sheet sheet = new Sheet()
                        {
                            Id = workbookPart.GetIdOfPart(worksheetPart),
                            SheetId = (uint)sheetCounter,
                            Name = sheetName
                        };
                        sheets.Append(sheet);

                        sheetCounter++;
                    }

                    workbookPart.Workbook.Save();
                }
                return memoryStream.ToArray();
            }
        }
    }
}
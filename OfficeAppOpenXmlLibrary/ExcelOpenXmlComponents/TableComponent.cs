using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml.Linq;
using DocumentFormat.OpenXml;

namespace OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents
{
	public class TableComponent
	{
		public static void AddTable(XElement table, Worksheet worksheet, out int rowCount, out int colCount, int startRow)
		{
			List<XElement> rows = table.Elements("row").ToList();
			rowCount = rows.Count;
			colCount = rows.Max(r => r.Elements("cell").Count());

			SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
			if (sheetData == null)
			{
				sheetData = new SheetData();
				worksheet.Append(sheetData);
			}

			for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
			{
				Row newRow = new Row() { RowIndex = (UInt32Value)(uint)(rowIndex + startRow) };
				List<XElement> cells = rows[rowIndex].Elements("cell").ToList();

				for (int colIndex = 0; colIndex < cells.Count; colIndex++)
				{
					string cellValue = cells[colIndex].Value;

					string? formula = cells[colIndex].Attribute("formula")?.Value;

					Cell cell = new Cell()
					{
						CellValue = new CellValue(cellValue),
						CellFormula = string.IsNullOrEmpty(formula) ? null : new CellFormula(formula),
					};

					if (double.TryParse(cellValue, out double numericValue))
					{
						cell.DataType = CellValues.Number;
					}
					else
					{
						cell.DataType = CellValues.String;
					}

					newRow.Append(cell);
				}

				sheetData.Append(newRow);
			}
		}
	}
}



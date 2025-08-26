using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Xml.Linq;

namespace OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents
{
    public class TableComponent
    {
        public static void AddTable(XElement table, Worksheet worksheet, out int rowCount, out int colCount)
        {
            string startCell = table.Attribute("startCell")?.Value ?? "A1";
            (int startRow, int startCol) = ParseA1(startCell); // 1-based

            var rows = table.Elements("row").ToList();
            rowCount = rows.Count;
            colCount = rows.Count == 0 ? 0 : rows.Max(r => r.Elements("cell").Count());

            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
            {
                sheetData = new SheetData();
                worksheet.Append(sheetData);
            }

            for (int r = 0; r < rows.Count; r++)
            {
                uint excelRowIndex = (uint)(startRow + r);

                // Var olan satır varsa al, yoksa oluştur
                var row = sheetData.Elements<Row>().FirstOrDefault(x => x.RowIndex == excelRowIndex);
                if (row == null)
                {
                    row = new Row() { RowIndex = excelRowIndex };
                    // Sıralı eklemek adına doğru yere eklemek isterseniz ek sıralama gerekebilir
                    sheetData.Append(row);
                }

                var cells = rows[r].Elements("cell").ToList();
                for (int c = 0; c < cells.Count; c++)
                {
                    int excelColIndex = startCol + c;
                    string cellRef = ToA1(excelRowIndex, excelColIndex);

                    string raw = cells[c].Value;
                    string? formula = cells[c].Attribute("formula")?.Value;

                    var cell = new Cell() { CellReference = cellRef };

                    if (!string.IsNullOrWhiteSpace(formula))
                    {
                        // '=' işaretini at
                        string f = formula.TrimStart('=');
                        cell.CellFormula = new CellFormula(f);

                        // 1) En kısa yol: XML'den 'cached' geliyorsa direkt yaz
                        string cached = cells[c].Attribute("cached")?.Value;
                        if (!string.IsNullOrWhiteSpace(cached))
                        {
                            cell.CellValue = new CellValue(cached); // ondalıklar '.' olmalı
                        }

                        // (İstersen buraya önceki "TryComputeCachedValue" mini hesaplayıcıyı da
                        // çağırıp cached yoksa otomatik hesaplatabilirsin.)
                    }
                    else if (double.TryParse(raw, System.Globalization.NumberStyles.Any,
                                             System.Globalization.CultureInfo.InvariantCulture,
                                             out var num))
                    {
                        // Sayı için DataType verme, yalnızca değer yaz
                        cell.CellValue = new CellValue(num.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        // Metin: InlineString
                        cell.DataType = CellValues.InlineString;
                        cell.InlineString = new InlineString(new Text(raw)
                        { Space = SpaceProcessingModeValues.Preserve });
                    }

                    // Var olan satırda aynı referansa sahip hücre varsa güncelle, yoksa ekle
                    var existing = row.Elements<Cell>().FirstOrDefault(x => x.CellReference?.Value == cellRef);
                    if (existing != null)
                    {
                        // mevcut hücreyi aynı pozisyonda yeni hücreyle değiştir
                        row.ReplaceChild(cell, existing); // veya row.ReplaceChild<Cell>(cell, existing);
                    }
                    else
                    {
                        row.Append(cell);
                    }

                }
            }
        }

        // "A1" -> (row=1, col=1)
        private static (int row, int col) ParseA1(string a1)
        {
            if (string.IsNullOrWhiteSpace(a1)) return (1, 1);
            int i = 0;
            int col = 0;
            while (i < a1.Length && char.IsLetter(a1[i]))
            {
                col = col * 26 + (char.ToUpperInvariant(a1[i]) - 'A' + 1);
                i++;
            }
            int row = 1;
            if (i < a1.Length && int.TryParse(a1.Substring(i), out int r)) row = r;
            return (row, col <= 0 ? 1 : col);
        }

        // (row, col) -> "A1"
        private static string ToA1(uint row, int col)
        {
            string letters = "";
            int c = col;
            while (c > 0)
            {
                int rem = (c - 1) % 26;
                letters = (char)('A' + rem) + letters;
                c = (c - 1) / 26;
            }
            return $"{letters}{row}";
        }
    }
}

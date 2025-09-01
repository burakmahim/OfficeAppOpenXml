using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents
{
    public static class TableComponent
    {
        // Grafiklerin tabloya bağlanabilmesi için global kayıt
        public static readonly Dictionary<string, TablePlacement> TableRegistry = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Tam sürüm: tabloyu ekler ve TablePlacement döner. (ExcelLibrary bu sürümü kullanır)
        /// </summary>
        public static TablePlacement AddTable(XElement table, Worksheet worksheet, string sheetName)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));

            string startCell = table.Attribute("startCell")?.Value ?? "A1";
            (int startRow, int startCol) = ParseA1(startCell); // 1-based

            var rows = table.Elements("row").ToList();
            int rowCount = rows.Count;
            int colCount = rows.Count == 0 ? 0 : rows.Max(r => r.Elements("cell").Count());

            // SheetData hazır mı?
            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
            {
                sheetData = new SheetData();
                worksheet.Append(sheetData);
            }

            for (int r = 0; r < rows.Count; r++)
            {
                uint excelRowIndex = (uint)(startRow + r);

                // Var olan satırı bul, yoksa oluştur
                var row = sheetData.Elements<Row>().FirstOrDefault(x => x.RowIndex == excelRowIndex);
                if (row == null)
                {
                    row = new Row { RowIndex = excelRowIndex };
                    sheetData.Append(row);
                }

                var cells = rows[r].Elements("cell").ToList();
                for (int c = 0; c < cells.Count; c++)
                {
                    int excelColIndex = startCol + c;
                    string cellRef = ToA1(excelRowIndex, excelColIndex);

                    var src = cells[c];
                    string raw = src.Value ?? string.Empty;
                    string? formula = src.Attribute("formula")?.Value;
                    string? cached = src.Attribute("cached")?.Value;

                    var cell = new Cell { CellReference = cellRef };

                    if (!string.IsNullOrWhiteSpace(formula))
                    {
                        // '=' işaretini kaldırarak formülü yaz
                        cell.CellFormula = new CellFormula(formula.TrimStart('='));
                        if (!string.IsNullOrWhiteSpace(cached))
                        {
                            // cached değer varsa CellValue olarak yaz (ondalık '.' olmalı)
                            cell.CellValue = new CellValue(cached);
                            // DataType'ı boş bırakıyoruz; Excel sayısal/ metin ayrımını cached içeriğine göre halleder
                        }
                    }
                    else if (double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out var num))
                    {
                        // Saf sayı
                        cell.CellValue = new CellValue(num.ToString(CultureInfo.InvariantCulture));
                        // DataType verilmez → sayısal
                    }
                    else
                    {
                        // Metin: InlineString kullan
                        cell.DataType = CellValues.InlineString;
                        cell.InlineString = new InlineString(new Text(raw) { Space = SpaceProcessingModeValues.Preserve });
                    }

                    // Aynı referansta önceden hücre varsa yerinde değiştir; yoksa ekle
                    var existing = row.Elements<Cell>().FirstOrDefault(x => string.Equals(x.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
                    if (existing != null)
                        row.ReplaceChild(cell, existing);
                    else
                        row.Append(cell);
                }
            }

            // Yerleşim hesapla
            var endRow = startRow + Math.Max(0, rowCount - 1);
            var endCol = startCol + Math.Max(0, colCount - 1);
            string rangeA1 = $"{startCell}:{ToA1((uint)endRow, endCol)}";

            var placement = new TablePlacement
            {
                SheetName = sheetName,
                StartRow = startRow,
                StartCol = startCol,
                RowCount = rowCount,
                ColCount = colCount,
                StartCell = startCell,
                RangeA1 = rangeA1
            };

            // İsteğe bağlı id → registry
            string? id = table.Attribute("id")?.Value;
            if (!string.IsNullOrWhiteSpace(id))
                TableRegistry[id] = placement;

            return placement;
        }

        /// <summary>
        /// Geriye dönük uyumluluk: sadece satır/sütun sayısını döndürür.
        /// </summary>
        public static void AddTable(XElement table, Worksheet worksheet, out int rowCount, out int colCount)
        {
            var placement = AddTable(table, worksheet, sheetName: "(unknown)");
            rowCount = placement.RowCount;
            colCount = placement.ColCount;
        }

        // ------------------------
        // Yardımcılar
        // ------------------------

        // "A1" -> (row=1, col=1) 1-based
        public static (int row, int col) ParseA1(string a1)
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

            return (row <= 0 ? 1 : row, col <= 0 ? 1 : col);
        }

        // (row, col) -> "A1"
        public static string ToA1(uint row, int col) => $"{ColumnIndexToLetters(col)}{row}";

        public static string ColumnIndexToLetters(int col)
        {
            if (col < 1) col = 1;
            string letters = "";
            int c = col;
            while (c > 0)
            {
                int rem = (c - 1) % 26;
                letters = (char)('A' + rem) + letters;
                c = (c - 1) / 26;
            }
            return letters;
        }

        public static int LettersToColumnIndex(string letters)
        {
            if (string.IsNullOrWhiteSpace(letters)) return 1;
            int col = 0;
            foreach (char ch in letters.Trim().ToUpperInvariant())
            {
                if (!char.IsLetter(ch)) break;
                col = col * 26 + (ch - 'A' + 1);
            }
            return Math.Max(1, col);
        }

        // Grafik/bağlama tarafının ihtiyaç duyduğu tablo yerleşimi bilgisi
        public struct TablePlacement
        {
            public string SheetName { get; set; }
            public int StartRow { get; set; }   // 1-based
            public int StartCol { get; set; }   // 1-based
            public int RowCount { get; set; }
            public int ColCount { get; set; }
            public string StartCell { get; set; } // örn "B3"
            public string RangeA1 { get; set; }   // örn "B3:D10"
        }
    }
}

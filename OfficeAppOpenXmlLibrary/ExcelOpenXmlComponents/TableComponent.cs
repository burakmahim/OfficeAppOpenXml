using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents
{
    /// <summary>
    /// XML tabloyu Worksheet'e yazar.
    /// - Her hücreye A1 stili CellReference atar.
    /// - Formül varsa CellFormula, yoksa CellValue + DataType (Number/String) set eder.
    /// - Dışarıya rowCount/colCount döner.
    /// Ekstralar:
    /// - startCol destekli AddTable aşırı yükleme
    /// - TablePlacement ile üst katmanda grafik aralığı üretimi
    /// - (Opsiyonel) CreateExcelTable ile gerçek Excel Table (ListObject) oluşturma
    /// </summary>
    public static class TableComponent
    {
        // =========================
        // Yardımcılar (A1 adresleme)
        // =========================

        /// <summary>1-bazlı kolonu harfe çevirir (1=A, 2=B, 27=AA).</summary>
        private static string ToColLetters(int colIndex1Based)
        {
            if (colIndex1Based <= 0)
                throw new ArgumentOutOfRangeException(nameof(colIndex1Based), "Column index must be >= 1");

            int dividend = colIndex1Based;
            string col = string.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                col = Convert.ToChar('A' + modulo) + col;
                dividend = (dividend - modulo) / 26;
            }
            return col;
        }

        /// <summary>A1 adresini döner. Örn: A1, $B$2.</summary>
        private static string A1(int row1Based, int col1Based, bool absolute = true)
        {
            if (row1Based <= 0) throw new ArgumentOutOfRangeException(nameof(row1Based));
            if (col1Based <= 0) throw new ArgumentOutOfRangeException(nameof(col1Based));

            string c = ToColLetters(col1Based);
            return absolute ? $"${c}${row1Based}" : $"{c}{row1Based}";
        }

        // =========================
        // Kamu API'leri
        // =========================

        /// <summary>
        /// Geriye dönük uyumluluk için: A sütunundan başlar (startCol=1).
        /// </summary>
        public static void AddTable(XElement table, Spreadsheet.Worksheet worksheet, out int rowCount, out int colCount, int startRow)
            => AddTable(table, worksheet, out rowCount, out colCount, startRow, startCol: 1);

        /// <summary>
        /// XML tablosunu belirtilen startRow/startCol'dan başlayarak yazar.
        /// Hücrelere A1 CellReference yazar; formül varsa CellFormula set eder.
        /// </summary>
        public static void AddTable(
            XElement table,
            Spreadsheet.Worksheet worksheet,
            out int rowCount,
            out int colCount,
            int startRow,
            int startCol)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
            if (startRow <= 0) throw new ArgumentOutOfRangeException(nameof(startRow));
            if (startCol <= 0) throw new ArgumentOutOfRangeException(nameof(startCol));

            var rows = table.Elements("row").ToList();
            rowCount = rows.Count;
            colCount = rows.Count == 0 ? 0 : rows.Max(r => r.Elements("cell").Count());

            var sheetData = worksheet.GetFirstChild<Spreadsheet.SheetData>();
            if (sheetData == null)
            {
                sheetData = new Spreadsheet.SheetData();
                worksheet.Append(sheetData);
            }

            // Satırları yaz
            for (int r = 0; r < rows.Count; r++)
            {
                uint excelRowIndex = (uint)(startRow + r);
                var newRow = new Spreadsheet.Row { RowIndex = excelRowIndex };

                var cells = rows[r].Elements("cell").ToList();

                for (int c = 0; c < cells.Count; c++)
                {
                    int excelColIndex = startCol + c;

                    string cellValue = cells[c].Value ?? string.Empty;
                    string formula = cells[c].Attribute("formula")?.Value;

                    var cell = new Spreadsheet.Cell
                    {
                        // A1 adresi (mutlak atamak şart değil; Excel işler)
                        CellReference = A1((int)excelRowIndex, excelColIndex, absolute: false)
                    };

                    if (!string.IsNullOrWhiteSpace(formula))
                    {
                        // Formül varsa değeri Excel hesaplar; CellValue eklemiyoruz
                        cell.CellFormula = new Spreadsheet.CellFormula(formula);
                    }
                    else
                    {
                        // Değer-set: sayı mı string mi?
                        if (double.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                        {
                            cell.DataType = Spreadsheet.CellValues.Number;
                            cell.CellValue = new Spreadsheet.CellValue(cellValue);
                        }
                        else
                        {
                            cell.DataType = Spreadsheet.CellValues.String;
                            cell.CellValue = new Spreadsheet.CellValue(cellValue);
                        }
                    }

                    newRow.Append(cell);
                }

                sheetData.Append(newRow);
            }
        }

        /// <summary>
        /// Üst katman için pratik: tabloyu yazar ve konum/boyut bilgilerini tek obje olarak döner.
        /// Bu sayede grafiğe aralık verirken placement.BuildRange(0/1/...) kullanabilirsin.
        /// </summary>
        public static TablePlacement AddTable(
            XElement table,
            Spreadsheet.Worksheet worksheet,
            string sheetName,
            int startRow,
            int startCol,
            out int rowCount,
            out int colCount)
        {
            AddTable(table, worksheet, out rowCount, out colCount, startRow, startCol);

            return new TablePlacement
            {
                SheetName = sheetName,
                StartRow = startRow,
                StartCol = startCol,
                RowCount = rowCount,
                ColCount = colCount
            };
        }

        // ==================================================
        // (OPSİYONEL) Gerçek Excel Table (ListObject) kurucu
        // ==================================================

        /// <summary>
        /// Worksheet üzerinde gerçek bir Excel Table (ListObject) oluşturur.
        /// Oluşan tablo A1 aralığını ve güvenli display name'i döner.
        /// </summary>
        public static (uint TableId, string ReferenceA1, string DisplayName) CreateExcelTable(
            WorksheetPart wsPart,
            string tableName,
            int startRow,
            int startCol,
            int rowCount,
            int colCount,
            IList<string> headerNames = null,
            string tableStyle = "TableStyleMedium2")
        {
            if (wsPart == null) throw new ArgumentNullException(nameof(wsPart));
            if (rowCount < 1 || colCount < 1) throw new ArgumentException("Table must have at least 1 row and 1 column.");
            if (startRow <= 0 || startCol <= 0) throw new ArgumentOutOfRangeException("startRow/startCol must be >= 1");

            var ws = wsPart.Worksheet;

            // TableParts düğümünü al/oluştur
            var tableParts = ws.GetFirstChild<Spreadsheet.TableParts>();
            if (tableParts == null)
            {
                tableParts = new Spreadsheet.TableParts() { Count = 0U };
                ws.Append(tableParts);
            }

            // Benzersiz TableId
            var existingIds = new HashSet<uint>();
            foreach (var tp in tableParts.Elements<Spreadsheet.TablePart>())
            {
                var part = wsPart.GetPartById(tp.Id) as TableDefinitionPart;
                if (part?.Table?.Id?.Value is uint idVal)
                    existingIds.Add(idVal);
            }
            uint nextId = 1;
            while (existingIds.Contains(nextId)) nextId++;

            // A1 referansı
            string startRef = A1(startRow, startCol);
            string endRef = A1(startRow + rowCount - 1, startCol + colCount - 1);
            string refA1 = $"{startRef}:{endRef}";

            // Güvenli tablo adı (DisplayName)
            string safeName = Regex.Replace(string.IsNullOrWhiteSpace(tableName) ? $"Table{nextId}" : tableName, @"[^\w]", "_");

            // Kolonlar
            var columns = new Spreadsheet.TableColumns() { Count = (uint)colCount };
            for (uint i = 0; i < colCount; i++)
            {
                string name = (headerNames != null && i < headerNames.Count && !string.IsNullOrWhiteSpace(headerNames[(int)i]))
                    ? headerNames[(int)i]
                    : $"Column{i + 1}";

                columns.Append(new Spreadsheet.TableColumn() { Id = i + 1, Name = name });
            }

            // TableDefinitionPart oluştur
            var tableDefPart = wsPart.AddNewPart<TableDefinitionPart>();
            string relId = wsPart.GetIdOfPart(tableDefPart);

            var table = new Spreadsheet.Table
            {
                Id = nextId,
                Name = safeName,
                DisplayName = safeName,
                Reference = refA1,
                AutoFilter = new Spreadsheet.AutoFilter() { Reference = refA1 },
                HeaderRowCount = 1U,
                TableColumns = columns,
                TableStyleInfo = new Spreadsheet.TableStyleInfo
                {
                    Name = tableStyle,
                    ShowFirstColumn = false,
                    ShowLastColumn = false,
                    ShowRowStripes = true,
                    ShowColumnStripes = false
                }
            };

            tableDefPart.Table = table;
            tableDefPart.Table.Save();

            // Worksheet'e TablePart referansı ekle
            tableParts.Append(new Spreadsheet.TablePart() { Id = relId });
            tableParts.Count = (uint)tableParts.Elements<Spreadsheet.TablePart>().Count();

            ws.Save();

            return (nextId, refA1, safeName);
        }

        // =========================
        // Yardımcı tip: Placement
        // =========================

        /// <summary>
        /// Tablo nereye yazıldı? Grafiğe aralık verirken kullanmak için.
        /// .NET Framework 4.8 uyumu için 'init' yerine 'set' kullanıldı.
        /// </summary>
        public sealed class TablePlacement
        {
            public string SheetName { get; set; } = "";
            public int StartRow { get; set; }
            public int StartCol { get; set; }
            public int RowCount { get; set; }
            public int ColCount { get; set; }

            /// <summary>
            /// Sadece veri satırlarını (header hariç) kapsayan A1 aralığı döndürür.
            /// colOffsetFromStart: 0 = ilk kolon, 1 = ikinci kolon...
            /// </summary>
            public string BuildRange(int colOffsetFromStart, bool absolute = true)
            {
                if (string.IsNullOrWhiteSpace(SheetName))
                    throw new InvalidOperationException("SheetName is required.");

                int firstDataRow = StartRow + 1;            // header'ı atla
                int lastDataRow = StartRow + RowCount - 1;  // verinin bittiği satır
                int col = StartCol + colOffsetFromStart;

                string a = A1(firstDataRow, col, absolute);
                string b = A1(lastDataRow, col, absolute);
                return $"'{SheetName}'!{a}:{b}";
            }
        }
    }
}

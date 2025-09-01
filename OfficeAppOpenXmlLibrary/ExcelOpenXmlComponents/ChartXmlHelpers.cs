using System;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents
{
    public static class ChartXmlHelpers
    {
        // ---- Tablodan referans alan sürüm ----
        // Örnek çağrı:
        // ChartXmlHelpers.BuildChartDefinitionFromTable(chartEl, placement, sheetName, "A", "B,C")
        public static ChartDefinition BuildChartDefinitionFromTable(
            XElement chartEl,
            TableComponent.TablePlacement placement,
            string sheetName,
            string categoryColumnLetter,
            string valueColumnsLettersCsv)
        {
            var def = new ChartDefinition
            {
                Type = ParseChartType((string?)chartEl.Attribute("type")),
                GroupingType = ParseGrouping((string?)chartEl.Attribute("grouping")),
                VaryColors = TryBool((string?)chartEl.Attribute("varyColors")),
                ShowCategoryAxis = TryBool((string?)chartEl.Attribute("showCategoryAxis")) ?? true,
                ShowValueAxis = TryBool((string?)chartEl.Attribute("showValueAxis")) ?? true,
                Title = TitleFromAttr(chartEl)
            };

            // Tablo yerleşimi
            var (startRow, startCol) = ParseA1(placement.StartCell); // 1-based
            int dataRows = Math.Max(0, placement.RowCount);
            int dataCols = Math.Max(0, placement.ColCount);

            // Kolon harfleri -> index
            int catColIndex = ColLetterToIndex(categoryColumnLetter); // 1-based
            var valueCols = (valueColumnsLettersCsv ?? "B")
                            .Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(ColLetterToIndex)
                            .Where(i => i > 0)
                            .ToList();

            // Sayfa adı tek tırnakla sarılsın (boşluk/özel karakterler için)
            string qSheet = $"'{sheetName}'";

            // Kategori aralığı: aynı satırlar, tek kolon
            string catColLetter = ColIndexToLetter(startCol + (catColIndex - 1));
            int firstDataRow = startRow;               // tablo ilk satırı veri kabul ediyorsan
            int lastDataRow = startRow + dataRows - 1;

            string catRange = $"{qSheet}!${catColLetter}${firstDataRow}:${catColLetter}${lastDataRow}";

            // Seriler: her value kolonu bir seri
            foreach (var valColIndex in valueCols)
            {
                string valColLetter = ColIndexToLetter(startCol + (valColIndex - 1));
                string valRange = $"{qSheet}!${valColLetter}${firstDataRow}:${valColLetter}${lastDataRow}";

                var series = new SeriesDefinition(def)
                {
                    Title = null // istersen chartEl içinden <series title="..."> da okuyup set edebilirsin
                };

                // Kategori: hücreye referans formülü
                var catRef = new C.StringReference
                {
                    Formula = new C.Formula(catRange)
                };

                // Değerler: hücreye referans formülü
                var valRef = new C.NumberReference
                {
                    Formula = new C.Formula(valRange)
                };

                // Series, ExcelChartSupport tarafında C.CategoryAxisData / C.Values içine eklenecek
                // Biz sadece SeriesDefinition.Points'ı boş bırakıyoruz; bağlama formüllerini ExcelChartSupport kuracaksa
                // buraya bir işaret koyalım:
                series.Points = null; // "tabloda" bağlanacağını ifade ediyoruz (literal cache yok)

                def.Series.Add(series);

                // Bu iki formülü ExcelChartSupport'ta kullanabilmek için SeriesDefinition içine geçirmenin iki yolu var:
                // 1) SeriesDefinition'a alan eklemek (Preferred)
                // 2) ExcelChartSupport.BuildBarOrLineSeries içinde tablo modunu tespit edip burada tekrar formül kurmak
                // Burada 1. yolu seçiyoruz:
                EnsureBindingHolders(series, catRef, valRef);
            }

            return def;
        }

        // ---- Inline veri (XML içinde <series><point .../></series>) ----
        public static ChartDefinition BuildChartDefinitionInline(XElement chartEl)
        {
            var def = new ChartDefinition
            {
                Type = ParseChartType((string?)chartEl.Attribute("type")),
                GroupingType = ParseGrouping((string?)chartEl.Attribute("grouping")),
                VaryColors = TryBool((string?)chartEl.Attribute("varyColors")),
                ShowCategoryAxis = TryBool((string?)chartEl.Attribute("showCategoryAxis")) ?? true,
                ShowValueAxis = TryBool((string?)chartEl.Attribute("showValueAxis")) ?? true,
                Title = TitleFromAttr(chartEl)
            };

            foreach (var sEl in chartEl.Elements("series"))
            {
                var s = new SeriesDefinition(def)
                {
                    Title = TitleFromAttr(sEl)
                };

                foreach (var pEl in sEl.Elements("point"))
                {
                    var p = new PointDefinition(s)
                    {
                        Category = (string?)pEl.Attribute("category"),
                        Value = TryDouble((string?)pEl.Attribute("value")),
                        X = TryDouble((string?)pEl.Attribute("x")),
                        Y = TryDouble((string?)pEl.Attribute("y"))
                    };
                    s.Points.Add(p);
                }

                def.Series.Add(s);
            }

            return def;
        }

        // ---------------- helpers ----------------

        private static ChartType ParseChartType(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return ChartType.Column;
            var norm = s.Trim().ToLowerInvariant();
            return norm switch
            {
                "bar" => ChartType.Bar,
                "column" => ChartType.Column,
                "line" => ChartType.Line,
                "area" => ChartType.Area,
                "pie" => ChartType.Pie,
                "doughnut" => ChartType.Doughnut,
                "scatter" => ChartType.Scatter,
                "bubble" => ChartType.Bubble,
                "radar" => ChartType.Radar,
                _ => ChartType.Column
            };
        }

        private static GroupingType? ParseGrouping(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            var norm = s.Trim().ToLowerInvariant();
            return norm switch
            {
                "clustered" => GroupingType.Clustered,
                "stacked" => GroupingType.Stacked,
                "percentstacked" or "percent_stacked" or "percent-stacked" => GroupingType.PercentStacked,
                "standard" => GroupingType.Standard,
                _ => null
            };
        }

        private static TitleDefinition? TitleFromAttr(XElement el)
        {
            var t = (string?)el.Attribute("title");
            return string.IsNullOrWhiteSpace(t) ? null : new TitleDefinition { Text = t };
        }

        private static bool? TryBool(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            if (bool.TryParse(s, out var b)) return b;
            if (s == "0") return false;
            if (s == "1") return true;
            return null;
        }

        private static double? TryDouble(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d)) return d;
            return null;
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
            if (i < a1.Length && int.TryParse(a1.Substring(i), out int r))
                row = r;
            return (row, col <= 0 ? 1 : col);
        }

        private static int ColLetterToIndex(string letter)
        {
            if (string.IsNullOrWhiteSpace(letter)) return 0;
            int col = 0;
            foreach (var ch in letter.Trim().ToUpperInvariant())
            {
                if (!char.IsLetter(ch)) break;
                col = col * 26 + (ch - 'A' + 1);
            }
            return col;
        }

        private static string ColIndexToLetter(int col)
        {
            if (col <= 0) return "A";
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

        // --- Seri içine binding taşıyıcıları eklemek için ufak hack ---
        // SeriesDefinition'a alan eklemeden, Tag property’lerini kullanıyoruz.
        private static void EnsureBindingHolders(SeriesDefinition s, C.StringReference catRef, C.NumberReference valRef)
        {
            // SeriesDefinition genişletemiyorsan, Tag’ler gibi bir sözlük yoksa:
            // en kolay yol: başlığa sentinel bas ve ExcelChartSupport’ta bunu yakalamak.
            // Ama daha temiz: SeriesDefinition'a iki opsiyonel alan ekle:
            //   public C.StringReference BoundCategories {get;set;}
            //   public C.NumberReference BoundValues {get;set;}
            // Aşağıdaki satırlar, bu alanları varsayarak yazıldı:
            try
            {
                var propCat = typeof(SeriesDefinition).GetProperty("BoundCategories");
                var propVal = typeof(SeriesDefinition).GetProperty("BoundValues");
                propCat?.SetValue(s, catRef);
                propVal?.SetValue(s, valRef);
            }
            catch { /* no-op */ }
        }
    }
}

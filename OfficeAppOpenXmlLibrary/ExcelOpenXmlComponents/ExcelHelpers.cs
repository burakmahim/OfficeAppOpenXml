// -----------------------------------------------------------------------------
// ChartXmlHelpers.cs
// ExcelLibrary.cs içinde kullanılan XML yardımcıları
// -----------------------------------------------------------------------------
using System;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents
{
    internal static class ExcelHelpers
    {
        // ---- Basit attribute okuyucular ----
        public static string Attr(XElement el, string name, string fallback = null)
            => el?.Attribute(name)?.Value ?? fallback;

        public static bool? BoolAttr(XElement el, string name)
        {
            var s = Attr(el, name);
            if (string.IsNullOrWhiteSpace(s)) return null;
            if (bool.TryParse(s, out var b)) return b;
            if (s == "1") return true;
            if (s == "0") return false;
            return null;
        }

        public static int? IntAttr(XElement el, string name)
        {
            var s = Attr(el, name);
            if (int.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v)) return v;
            return null;
        }

        public static double? DoubleAttr(XElement el, string name)
        {
            var s = Attr(el, name);
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v)) return v;
            return null;
        }

        /// <summary>
        /// "clustered", "percent_stacked", "Percent-Stacked" gibi yazımları enum değerine çevirir.
        /// </summary>
        public static TEnum? EnumAttr<TEnum>(XElement el, string name) where TEnum : struct
        {
            var s = Attr(el, name);
            if (string.IsNullOrWhiteSpace(s)) return default;

            string norm = s.Replace("-", "").Replace("_", "").Trim().ToLowerInvariant();

            foreach (var candidate in Enum.GetNames(typeof(TEnum)))
            {
                var candNorm = candidate.Replace("_", "").ToLowerInvariant();
                if (candNorm == norm && Enum.TryParse<TEnum>(candidate, true, out var parsed))
                    return (TEnum)parsed;
            }

            if (Enum.TryParse<TEnum>(s, true, out var parsed2))
                return parsed2;

            return default;
        }

        // ---- A1 adres yardımcıları ----
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
            return (row, col <= 0 ? 1 : col);
        }

        public static string ToA1(int row, int col)
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

        // ---- Tablo referansı çözümlemeleri ----
        // chart xml: <chart tableRef="Table1" categoryCol="1" valueCol="2" headerRow="true" />
        // veya:     <chart tableRef="Table1" catRange="B2:B10" valRange="C2:C10" />
        public static bool TryGetTableAndRanges(
            XElement chartEl,
            out string tableName,
            out string catRange,
            out string valRange)
        {
            tableName = Attr(chartEl, "tableRef");
            catRange = Attr(chartEl, "catRange");
            valRange = Attr(chartEl, "valRange");

            if (string.IsNullOrWhiteSpace(tableName))
                return false;

            // Kolon indeksleri verilmişse (1-based), bunları aralık stringine çevirmeyi
            // ExcelLibrary tarafında SheetStartCell + TableRegistry ile yapıyoruz.
            return true;
        }

        // Serilerin doğrudan içte tanımlandığı senaryo:
        // <chart> <series title="S1"><point category="A" value="10" /></series> ... </chart>
        public static bool HasInlineSeries(XElement chartEl)
            => chartEl.Elements("series").Any();

        // Güvenli sayı çeviri (virgül nokta farklarına takılmasın)
        public static string ToInvariantString(double v)
            => v.ToString(CultureInfo.InvariantCulture);
    }
}

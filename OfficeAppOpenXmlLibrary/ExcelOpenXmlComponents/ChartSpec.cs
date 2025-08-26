// ChartSpec.cs
using System.Collections.Generic;

namespace OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents
{
    public class ChartSpec
    {
        public string Type { get; set; } = "column";
        public string? Title { get; set; }

        // A1 referans mod (eski davranış)
        public string CategoryRange { get; set; } = "'Sheet1'!$A$2:$A$6";
        public List<string> ValueRanges { get; set; } = new();
        public List<string>? SeriesNames { get; set; }

        // === INLINE VERİ (yeni) ===
        // Kategoriler: ["Ocak","Şubat",...]
        public List<string>? CategoriesInline { get; set; }
        // Seriler: her seri için sayı listesi (kategorilerle aynı uzunlukta)
        public List<List<double?>>? ValuesInline { get; set; }

        public string? FromCell { get; set; } = "D2";
        public string? ToCell { get; set; } = "L20";

        // (opsiyonel diğer ayarlar…)
        public bool? ShowLegend { get; set; }
        public string? LegendPosition { get; set; }
        public bool? PlotVisibleOnly { get; set; }
        public string? BlanksAs { get; set; }
        public bool? ShowDataLabelsOverMaximum { get; set; }
        public bool? VaryColors { get; set; }
        public string? Grouping { get; set; }
        public int? GapWidth { get; set; }
        public int? Overlap { get; set; }
        public bool? RoundedCorners { get; set; }
        public bool? ShowDataLabels { get; set; }
        public string? DataLabelPos { get; set; }
        public bool? ShowLabelValue { get; set; }
        public bool? ShowLabelPerc { get; set; }
        public bool? ShowCategoryAxis { get; set; }
        public bool? ShowValueAxis { get; set; }
        public string? CategoryAxisTitle { get; set; }
        public string? ValueAxisTitle { get; set; }
        public string? MarkerSymbol { get; set; }
        public int? MarkerSize { get; set; }
        public int? FirstSliceAngle { get; set; }
        public int? DoughnutHoleSize { get; set; }
        public string? ScatterStyle { get; set; }
        public string? RadarStyle { get; set; }
        public int? BubbleScale { get; set; }
        public bool? Bubble3D { get; set; }
        public List<string>? BubbleSizeRanges { get; set; }
    }
}

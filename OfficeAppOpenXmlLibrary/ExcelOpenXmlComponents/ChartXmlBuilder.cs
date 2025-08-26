using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;          // sadece DEBUG doğrulaması için
using DocumentFormat.OpenXml.Validation;         // sadece DEBUG doğrulaması için
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents
{
    public static class ChartXmlBuilder
    {
        private static uint _axisSeed = 48650000;
        private static uint NextAxisId() => ++_axisSeed;

        public static C.ChartSpace BuildChartSpace(ChartSpec s)
        {
            if (s == null) throw new ArgumentNullException(nameof(s));

            bool hasInline = s.ValuesInline != null && s.ValuesInline.Count > 0;
            bool hasRef = s.ValueRanges != null && s.ValueRanges.Count > 0;
            if (!hasInline && !hasRef)
                throw new ArgumentException("Grafik için ya ValuesInline ya da ValueRanges dolu olmalı.");

            var chartSpace = new C.ChartSpace();
            chartSpace.Append(new C.EditingLanguage { Val = "en-US" });
            if (s.RoundedCorners.HasValue)
                chartSpace.Append(new C.RoundedCorners { Val = s.RoundedCorners.Value });

            var chart = new C.Chart();

            // Başlık
            if (!string.IsNullOrWhiteSpace(s.Title))
            {
                chart.Append(new C.Title(
                    new C.ChartText(new C.RichText(
                        new A.BodyProperties(), new A.ListStyle(),
                        new A.Paragraph(new A.Run(new A.Text(s.Title)))
                    )),
                    new C.Layout()
                ));
            }
            else
            {
                chart.Append(new C.AutoTitleDeleted { Val = true });
            }

            // PlotArea
            var plotArea = new C.PlotArea();
            plotArea.Append(new C.Layout());

            switch ((s.Type ?? "column").Trim().ToLowerInvariant())
            {
                case "column": BuildBarLike(plotArea, s, isColumn: true); AddAxes(plotArea, s); break;
                case "bar": BuildBarLike(plotArea, s, isColumn: false); AddAxes(plotArea, s); break;
                case "line": BuildLine(plotArea, s); AddAxes(plotArea, s); break;
                case "area": BuildArea(plotArea, s); AddAxes(plotArea, s); break;
                case "pie": BuildPie(plotArea, s); break;
                case "doughnut": BuildDoughnut(plotArea, s); break;
                case "radar": BuildRadar(plotArea, s); AddAxes(plotArea, s); break;
                case "scatter": BuildScatter(plotArea, s); AddAxes(plotArea, s); break;
                case "bubble": BuildBubble(plotArea, s); AddAxes(plotArea, s); break;
                default: throw new NotSupportedException("Desteklenmeyen grafik tipi: " + s.Type);
            }

            chart.Append(plotArea);

            // Legend
            if (s.ShowLegend == true)
                chart.Append(new C.Legend(
                    new C.LegendPosition { Val = MapLegendPos(s.LegendPosition) },
                    new C.Layout()
                ));

            // Görünürlük / boşluk gösterimi / aşımda etiket
            if (s.PlotVisibleOnly.HasValue)
                chart.Append(new C.PlotVisibleOnly { Val = s.PlotVisibleOnly.Value });

            if (!string.IsNullOrEmpty(s.BlanksAs))
                chart.Append(new C.DisplayBlanksAs { Val = MapBlanksAs(s.BlanksAs) });

            if (s.ShowDataLabelsOverMaximum.HasValue)
                chart.Append(new C.ShowDataLabelsOverMaximum { Val = s.ShowDataLabelsOverMaximum.Value });

            chartSpace.Append(chart);
            return chartSpace;
        }

        // ============================
        //  CHART BUILDERS
        // ============================

        // Bar / Column
        private static void BuildBarLike(C.PlotArea pa, ChartSpec s, bool isColumn)
        {
            var ch = pa.AppendChild(new C.BarChart(
                new C.BarDirection { Val = isColumn ? C.BarDirectionValues.Column : C.BarDirectionValues.Bar },
                new C.BarGrouping { Val = MapGrouping(s.Grouping) },
                new C.VaryColors { Val = s.VaryColors ?? false }
            ));

            int seriesCount = s.ValuesInline?.Count ?? s.ValueRanges.Count;

            for (int i = 0; i < seriesCount; i++)
            {
                var ser = ch.AppendChild(new C.BarChartSeries(
                    new C.Index { Val = (uint)i },
                    new C.Order { Val = (uint)i }
                ));

                string name = (s.SeriesNames != null && i < s.SeriesNames.Count) ? s.SeriesNames[i] : $"Series {i + 1}";
                ser.Append(new C.SeriesText(new C.NumericValue { Text = name }));

                ser.Append(BuildCatData(s));
                ser.Append(BuildValData(s, i));
            }

            // Şema sırası: ser* -> dLbls? -> gapWidth? -> overlap? -> axId{2}
            AddDataLabelsIfAny(ch, s);

            if (s.GapWidth.HasValue)
                ch.Append(new C.GapWidth { Val = (ushort)Math.Max(0, Math.Min(500, s.GapWidth.Value)) });

            if (s.Overlap.HasValue)
                ch.Append(new C.Overlap { Val = (sbyte)Math.Max(-100, Math.Min(100, s.Overlap.Value)) });
        }

        // Line
        private static void BuildLine(C.PlotArea pa, ChartSpec s)
        {
            var ch = pa.AppendChild(new C.LineChart(
                new C.Grouping { Val = C.GroupingValues.Standard },
                new C.VaryColors { Val = s.VaryColors ?? false }
            ));

            int seriesCount = s.ValuesInline?.Count ?? s.ValueRanges.Count;

            for (int i = 0; i < seriesCount; i++)
            {
                var ser = ch.AppendChild(new C.LineChartSeries(
                    new C.Index { Val = (uint)i },
                    new C.Order { Val = (uint)i }
                ));

                string name = (s.SeriesNames != null && i < s.SeriesNames.Count) ? s.SeriesNames[i] : $"Series {i + 1}";
                ser.Append(new C.SeriesText(new C.NumericValue { Text = name }));

                var marker = BuildMarker(s);
                if (marker != null) ser.Append(marker);

                ser.Append(BuildCatData(s));
                ser.Append(BuildValData(s, i));
            }

            AddDataLabelsIfAny(ch, s);
        }

        // Area
        private static void BuildArea(C.PlotArea pa, ChartSpec s)
        {
            var ch = pa.AppendChild(new C.AreaChart(
                new C.Grouping { Val = C.GroupingValues.Standard },
                new C.VaryColors { Val = s.VaryColors ?? false }
            ));

            int seriesCount = s.ValuesInline?.Count ?? s.ValueRanges.Count;

            for (int i = 0; i < seriesCount; i++)
            {
                var ser = ch.AppendChild(new C.AreaChartSeries(
                    new C.Index { Val = (uint)i },
                    new C.Order { Val = (uint)i }
                ));

                string name = (s.SeriesNames != null && i < s.SeriesNames.Count) ? s.SeriesNames[i] : $"Series {i + 1}";
                ser.Append(new C.SeriesText(new C.NumericValue { Text = name }));

                ser.Append(BuildCatData(s));
                ser.Append(BuildValData(s, i));
            }

            AddDataLabelsIfAny(ch, s);
        }

        // Pie
        private static void BuildPie(C.PlotArea pa, ChartSpec s)
        {
            var ch = pa.AppendChild(new C.PieChart(
                new C.VaryColors { Val = s.VaryColors ?? true }
            ));

            int seriesCount = s.ValuesInline?.Count ?? s.ValueRanges.Count;

            for (int i = 0; i < seriesCount; i++)
            {
                var ser = ch.AppendChild(new C.PieChartSeries(
                    new C.Index { Val = (uint)i },
                    new C.Order { Val = (uint)i }
                ));

                string name = (s.SeriesNames != null && i < s.SeriesNames.Count) ? s.SeriesNames[i] : $"Series {i + 1}";
                ser.Append(new C.SeriesText(new C.NumericValue { Text = name }));

                ser.Append(BuildCatData(s));
                ser.Append(BuildValData(s, i));
            }

            // Şema sırası: varyColors? -> ser* -> dLbls? -> firstSliceAng?  (pie’de holeSize yok)
            AddDataLabelsIfAny(ch, s, isPieLike: true);

            if (s.FirstSliceAngle.HasValue)
                ch.Append(new C.FirstSliceAngle { Val = (ushort)Math.Max(0, Math.Min(360, s.FirstSliceAngle.Value)) });
        }

        // Doughnut
        private static void BuildDoughnut(C.PlotArea pa, ChartSpec s)
        {
            var ch = pa.AppendChild(new C.DoughnutChart(
                new C.VaryColors { Val = s.VaryColors ?? true }
            ));

            for (int i = 0; i < (s.ValuesInline?.Count ?? s.ValueRanges.Count); i++)
            {
                var ser = ch.AppendChild(new C.PieChartSeries(
                    new C.Index { Val = (uint)i },
                    new C.Order { Val = (uint)i }
                ));

                ser.Append(new C.SeriesText(new C.NumericValue { Text = s.SeriesNames?[i] ?? $"Series {i + 1}" }));
                ser.Append(BuildCatData(s));
                ser.Append(BuildValData(s, i));
            }

            AddDataLabelsIfAny(ch, s, isPieLike: true);

            // Mutlaka ekle:
            ch.Append(new C.FirstSliceAngle { Val = (ushort)(s.FirstSliceAngle ?? 0) });
            ch.Append(new C.HoleSize { Val = (byte)(s.DoughnutHoleSize ?? 50) });
        }


        // Scatter
        private static void BuildScatter(C.PlotArea pa, ChartSpec s)
        {
            var ch = pa.AppendChild(new C.ScatterChart(
                new C.ScatterStyle { Val = MapScatterStyle(s.ScatterStyle) },
                new C.VaryColors { Val = s.VaryColors ?? false }
            ));

            int seriesCount = s.ValuesInline?.Count ?? s.ValueRanges?.Count ?? 0;
            if (seriesCount == 0)
                throw new InvalidOperationException("Scatter için en az bir seri gerekli.");

            for (int i = 0; i < seriesCount; i++)
            {
                var ser = ch.AppendChild(new C.ScatterChartSeries(
                    new C.Index { Val = (uint)i },
                    new C.Order { Val = (uint)i }
                ));

                ser.Append(new C.SeriesText(new C.NumericValue { Text = s.SeriesNames?[i] ?? $"Series {i + 1}" }));

                // DOĞRU SIRA: marker → xVal → yVal
                var marker = BuildMarker(s);
                if (marker != null) ser.Append(marker);

                ser.Append(BuildXValuesForScatterBubble(s));
                ser.Append(BuildYValues(s, i));
            }

            AddDataLabelsIfAny(ch, s);
        }


        // Radar
        private static void BuildRadar(C.PlotArea pa, ChartSpec s)
        {
            var ch = pa.AppendChild(new C.RadarChart(
                new C.RadarStyle { Val = MapRadarStyle(s.RadarStyle) },
                new C.VaryColors { Val = s.VaryColors ?? false }
            ));

            int seriesCount = s.ValuesInline?.Count ?? s.ValueRanges.Count;

            for (int i = 0; i < seriesCount; i++)
            {
                var ser = ch.AppendChild(new C.RadarChartSeries(
                    new C.Index { Val = (uint)i },
                    new C.Order { Val = (uint)i }
                ));

                string name = (s.SeriesNames != null && i < s.SeriesNames.Count) ? s.SeriesNames[i] : $"Series {i + 1}";
                ser.Append(new C.SeriesText(new C.NumericValue { Text = name }));

                ser.Append(BuildCatData(s));
                ser.Append(BuildValData(s, i));
            }

            AddDataLabelsIfAny(ch, s);
        }

        // Bubble
        private static void BuildBubble(C.PlotArea pa, ChartSpec s)
        {
            var ch = pa.AppendChild(new C.BubbleChart(
                new C.VaryColors { Val = s.VaryColors ?? false }
            ));

            int seriesCount = s.ValuesInline?.Count ?? s.ValueRanges?.Count ?? 0;
            if (seriesCount == 0)
                throw new InvalidOperationException("Bubble için en az bir seri gerekli (ValuesInline veya ValueRanges).");

            for (int i = 0; i < seriesCount; i++)
            {
                var ser = ch.AppendChild(new C.BubbleChartSeries(
                    new C.Index { Val = (uint)i },
                    new C.Order { Val = (uint)i }
                ));

                string name = (s.SeriesNames != null && i < s.SeriesNames.Count) ? s.SeriesNames[i] : $"Series {i + 1}";
                ser.Append(new C.SeriesText(new C.NumericValue { Text = name }));

                // X ve Y
                ser.Append(BuildXValuesForScatterBubble(s));
                ser.Append(BuildYValues(s, i));

                if (s.BubbleSizeRanges != null && i < s.BubbleSizeRanges.Count && !string.IsNullOrWhiteSpace(s.BubbleSizeRanges[i]))
                    ser.Append(new C.BubbleSize(new C.NumberReference(new C.Formula(s.BubbleSizeRanges[i]))));
            }

            AddDataLabelsIfAny(ch, s);

            if (s.Bubble3D.HasValue) ch.Append(new C.Bubble3D { Val = s.Bubble3D.Value });
            if (s.BubbleScale.HasValue) ch.Append(new C.BubbleScale { Val = (ushort)Math.Max(0, Math.Min(300, s.BubbleScale.Value)) });
        }

        // ============================
        //  AXES / LABELS / MARKER
        // ============================

        private static void AddAxes(C.PlotArea pa, ChartSpec s)
        {
            // Scatter/Bubble: 2 adet ValueAxis
            bool hasScatterOrBubble = pa.Elements<C.ScatterChart>().Any() || pa.Elements<C.BubbleChart>().Any();
            if (hasScatterOrBubble)
            {
                uint xId = NextAxisId(), yId = NextAxisId();

                var xVal = new C.ValueAxis(
                    new C.AxisId { Val = xId },
                    new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                    new C.Delete { Val = false },
                    new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
                    new C.MajorGridlines(),
                    new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
                    new C.CrossingAxis { Val = yId },
                    new C.Crosses { Val = C.CrossesValues.AutoZero }
                );

                var yVal = new C.ValueAxis(
                    new C.AxisId { Val = yId },
                    new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                    new C.Delete { Val = false },
                    new C.AxisPosition { Val = C.AxisPositionValues.Left },
                    new C.MajorGridlines(),
                    new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
                    new C.CrossingAxis { Val = xId },
                    new C.Crosses { Val = C.CrossesValues.AutoZero },
                    new C.CrossBetween { Val = C.CrossBetweenValues.Between }
                );

                pa.Append(xVal);
                pa.Append(yVal);

                foreach (var sc in pa.Elements<C.ScatterChart>())
                { sc.Append(new C.AxisId { Val = xId }); sc.Append(new C.AxisId { Val = yId }); }

                foreach (var bb in pa.Elements<C.BubbleChart>())
                { bb.Append(new C.AxisId { Val = xId }); bb.Append(new C.AxisId { Val = yId }); }

                return;
            }

            // Bar/Line/Area/Radar: CategoryAxis + ValueAxis
            uint cat = NextAxisId(), val = NextAxisId();

            var catAxis = new C.CategoryAxis(
                new C.AxisId { Val = cat },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = !(s.ShowCategoryAxis ?? true) },
                new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
                new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
                new C.CrossingAxis { Val = val },
                new C.Crosses { Val = C.CrossesValues.AutoZero },
                new C.AutoLabeled { Val = true },
                new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
                new C.LabelOffset { Val = 100 }
            );
            if (!string.IsNullOrWhiteSpace(s.CategoryAxisTitle))
                catAxis.Append(new C.Title(new C.ChartText(new C.RichText(
                    new A.BodyProperties(), new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(s.CategoryAxisTitle)))
                ))));
            pa.Append(catAxis);

            var valAxis = new C.ValueAxis(
                new C.AxisId { Val = val },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = !(s.ShowValueAxis ?? true) },
                new C.AxisPosition { Val = C.AxisPositionValues.Left },
                new C.MajorGridlines(),
                new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
                new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
                new C.CrossingAxis { Val = cat },
                new C.Crosses { Val = C.CrossesValues.AutoZero },
                new C.CrossBetween { Val = C.CrossBetweenValues.Between }
            );
            if (!string.IsNullOrWhiteSpace(s.ValueAxisTitle))
                valAxis.Append(new C.Title(new C.ChartText(new C.RichText(
                    new A.BodyProperties(), new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(s.ValueAxisTitle)))
                ))));
            pa.Append(valAxis);

            foreach (var bc in pa.Elements<C.BarChart>())
            { bc.Append(new C.AxisId { Val = cat }); bc.Append(new C.AxisId { Val = val }); }

            foreach (var lc in pa.Elements<C.LineChart>())
            { lc.Append(new C.AxisId { Val = cat }); lc.Append(new C.AxisId { Val = val }); }

            foreach (var ac in pa.Elements<C.AreaChart>())
            { ac.Append(new C.AxisId { Val = cat }); ac.Append(new C.AxisId { Val = val }); }

            foreach (var rd in pa.Elements<C.RadarChart>())
            { rd.Append(new C.AxisId { Val = cat }); rd.Append(new C.AxisId { Val = val }); }
        }

        private static void AddDataLabelsIfAny(OpenXmlCompositeElement chart, ChartSpec s, bool isPieLike = false)
        {
            if (s.ShowDataLabels != true && s.ShowLabelValue != true && (!isPieLike || s.ShowLabelPerc != true))
                return;

            var dl = new C.DataLabels();

            if (!string.IsNullOrWhiteSpace(s.DataLabelPos))
                dl.Append(new C.DataLabelPosition { Val = MapDLPos(s.DataLabelPos, isPieLike) });

            if (s.ShowLabelValue == true) dl.Append(new C.ShowValue { Val = true });
            if (isPieLike && s.ShowLabelPerc == true) dl.Append(new C.ShowPercent { Val = true });

            chart.Append(dl);
        }

        private static C.Marker? BuildMarker(ChartSpec s)
        {
            if (string.IsNullOrWhiteSpace(s.MarkerSymbol) && !s.MarkerSize.HasValue)
                return null;

            var m = new C.Marker();

            if (!string.IsNullOrWhiteSpace(s.MarkerSymbol))
                m.Append(new C.Symbol { Val = MapMarker(s.MarkerSymbol) });

            if (s.MarkerSize.HasValue)
                m.Append(new C.Size { Val = (byte)Math.Max(2, Math.Min(72, s.MarkerSize.Value)) });

            return m;
        }

        // ============================
        //  HELPERS (inline vs ref)
        // ============================

        // Kategoriler: inline (strLit) ya da A1 referansı (strRef)
        private static C.CategoryAxisData BuildCatData(ChartSpec s)
        {
            if (s.CategoriesInline != null && s.CategoriesInline.Count > 0)
            {
                var lit = new C.StringLiteral();
                lit.Append(new C.PointCount { Val = (uint)s.CategoriesInline.Count });
                for (int i = 0; i < s.CategoriesInline.Count; i++)
                {
                    lit.AppendChild(new C.StringPoint
                    {
                        Index = (uint)i,
                        NumericValue = new C.NumericValue(s.CategoriesInline[i] ?? string.Empty)
                    });
                }
                return new C.CategoryAxisData(lit);
            }
            return new C.CategoryAxisData(new C.StringReference(new C.Formula(s.CategoryRange)));
        }

        // Değerler (çoğu grafik): inline (numLit) ya da A1 referansı (numRef)
        private static C.Values BuildValData(ChartSpec s, int seriesIndex)
        {
            if (s.ValuesInline != null &&
                seriesIndex < s.ValuesInline.Count &&
                s.ValuesInline[seriesIndex] != null)
            {
                var vals = s.ValuesInline[seriesIndex];
                var lit = new C.NumberLiteral();
                lit.Append(new C.FormatCode { Text = "General" });
                lit.Append(new C.PointCount { Val = (uint)vals.Count });

                for (int i = 0; i < vals.Count; i++)
                {
                    var v = vals[i]?.ToString(CultureInfo.InvariantCulture) ?? "0";
                    lit.AppendChild(new C.NumericPoint
                    {
                        Index = (uint)i,
                        NumericValue = new C.NumericValue(v)
                    });
                }
                return new C.Values(lit);
            }
            return new C.Values(new C.NumberReference(new C.Formula(s.ValueRanges[seriesIndex])));
        }

        // Y değerleri (Scatter/Bubble için)
        private static C.YValues BuildYValues(ChartSpec s, int seriesIndex)
        {
            if (s.ValuesInline != null &&
                seriesIndex < s.ValuesInline.Count &&
                s.ValuesInline[seriesIndex] != null)
            {
                var vals = s.ValuesInline[seriesIndex];
                var lit = new C.NumberLiteral();
                lit.Append(new C.FormatCode { Text = "General" });
                lit.Append(new C.PointCount { Val = (uint)vals.Count });

                for (int i = 0; i < vals.Count; i++)
                {
                    var v = vals[i]?.ToString(CultureInfo.InvariantCulture) ?? "0";
                    lit.AppendChild(new C.NumericPoint
                    {
                        Index = (uint)i,
                        NumericValue = new C.NumericValue(v)
                    });
                }
                return new C.YValues(lit);
            }
            return new C.YValues(new C.NumberReference(new C.Formula(s.ValueRanges[seriesIndex])));
        }

        // ============================
        //  MAPPERS
        // ============================

        private static C.BarGroupingValues MapGrouping(string? g) =>
            ((g ?? "clustered").Trim().ToLowerInvariant()) switch
            {
                "standard" => C.BarGroupingValues.Standard,
                "stacked" => C.BarGroupingValues.Stacked,
                "percentstacked" => C.BarGroupingValues.PercentStacked,
                _ => C.BarGroupingValues.Clustered
            };

        private static C.LegendPositionValues MapLegendPos(string? p) =>
            ((p ?? "right").Trim().ToLowerInvariant()) switch
            {
                "left" => C.LegendPositionValues.Left,
                "top" => C.LegendPositionValues.Top,
                "bottom" => C.LegendPositionValues.Bottom,
                "corner" => C.LegendPositionValues.TopRight,
                _ => C.LegendPositionValues.Right
            };

        private static C.DisplayBlanksAsValues MapBlanksAs(string? b) =>
            ((b ?? "gap").Trim().ToLowerInvariant()) switch
            {
                "zero" => C.DisplayBlanksAsValues.Zero,
                "span" => C.DisplayBlanksAsValues.Span,
                _ => C.DisplayBlanksAsValues.Gap
            };

        private static C.DataLabelPositionValues MapDLPos(string pos, bool pie)
        {
            pos = pos.Trim().ToLowerInvariant();
            if (pie)
            {
                if (pos == "bestfit") return C.DataLabelPositionValues.BestFit;
                if (pos == "outsideend") return C.DataLabelPositionValues.OutsideEnd;
            }
            return pos switch
            {
                "center" => C.DataLabelPositionValues.Center,
                "insideend" => C.DataLabelPositionValues.InsideEnd,
                "insidebase" => C.DataLabelPositionValues.InsideBase,
                "outsideend" => C.DataLabelPositionValues.OutsideEnd,
                "left" => C.DataLabelPositionValues.Left,
                "right" => C.DataLabelPositionValues.Right,
                "top" => C.DataLabelPositionValues.Top,
                "bottom" => C.DataLabelPositionValues.Bottom,
                _ => C.DataLabelPositionValues.Center
            };
        }

        private static C.MarkerStyleValues MapMarker(string sym) =>
            ((sym ?? "circle").Trim().ToLowerInvariant()) switch
            {
                "square" => C.MarkerStyleValues.Square,
                "triangle" => C.MarkerStyleValues.Triangle,
                "x" => C.MarkerStyleValues.X,
                "diamond" => C.MarkerStyleValues.Diamond,
                "none" => C.MarkerStyleValues.None,
                _ => C.MarkerStyleValues.Circle
            };

        private static C.ScatterStyleValues MapScatterStyle(string? s) =>
            ((s ?? "marker").Trim().ToLowerInvariant()) switch
            {
                "line" => C.ScatterStyleValues.Line,
                "linemarker" => C.ScatterStyleValues.LineMarker,
                "smooth" => C.ScatterStyleValues.Smooth,
                "smoothmarker" => C.ScatterStyleValues.SmoothMarker,
                _ => C.ScatterStyleValues.Marker
            };

        private static C.RadarStyleValues MapRadarStyle(string? s) =>
            ((s ?? "standard").Trim().ToLowerInvariant()) switch
            {
                "marker" => C.RadarStyleValues.Marker,
                "filled" => C.RadarStyleValues.Filled,
                _ => C.RadarStyleValues.Standard
            };

        // Scatter/Bubble X değerleri: NumberRef (CategoryRange) veya numLit (CategoriesInline sayısal ise)
        private static C.XValues BuildXValuesForScatterBubble(ChartSpec s)
        {
            var x = new C.XValues();

            if (!string.IsNullOrWhiteSpace(s.CategoryRange))
            {
                x.Append(new C.NumberReference(new C.Formula(s.CategoryRange)));
                return x;
            }

            // CategoriesInline varsa ve tamamı sayısalsa numLit yaz
            if (s.CategoriesInline != null && s.CategoriesInline.Count > 0)
            {
                bool allNumeric = true;
                var lit = new C.NumberLiteral();
                lit.Append(new C.FormatCode { Text = "General" });
                lit.Append(new C.PointCount { Val = (uint)s.CategoriesInline.Count });

                for (int i = 0; i < s.CategoriesInline.Count; i++)
                {
                    var t = s.CategoriesInline[i] ?? "";
                    if (!double.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                    {
                        allNumeric = false;
                        break;
                    }
                    lit.AppendChild(new C.NumericPoint
                    {
                        Index = (uint)i,
                        NumericValue = new C.NumericValue(t)
                    });
                }

                if (allNumeric)
                {
                    x.Append(lit);
                    return x;
                }
            }

            throw new InvalidOperationException(
                "Scatter/Bubble için X değerleri gerekli. 'CategoryRange' sayısal bir aralık verin ya da 'CategoriesInline' değerleri sayısal olsun.");
        }


#if DEBUG
        // İstersen Excel üretimi bittikten sonra çağır:
        // ValidateCharts(spreadsheetDocument);
        private static void ValidateCharts(SpreadsheetDocument doc)
        {
            var validator = new OpenXmlValidator(FileFormatVersions.Office2013);
            foreach (var wsPart in doc.WorkbookPart.WorksheetParts)
            {
                var dp = wsPart.DrawingsPart;
                if (dp == null) continue;

                foreach (var cp in dp.ChartParts)
                {
                    var errs = validator.Validate(cp.ChartSpace);
                    foreach (var err in errs)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[Chart VALIDATION] Part: {err.Part?.Uri} | Path: {err.Path?.XPath} | Desc: {err.Description}");
                    }
                }
            }
        }
#endif


    }
}

// -----------------------------------------------------------------------------
// ExcelChartSupport.cs
// Excel çalışma sayfasına (WorksheetPart) ChartDefinition kullanarak grafik ekler
// - OpenXML SDK: DocumentFormat.OpenXml (Spreadsheet + Drawing + Drawing.Charts)
// - Grafik içeriğini cache (literal) olarak yazar; hücre aralığına bağlamaz.
// -----------------------------------------------------------------------------

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents
{
    public static class ExcelChartSupport
    {
        /// <summary>
        /// Belirtilen çalışma sayfasına grafiği ekler ve verilen hücre aralığına (fromCell→toCell) yerleştirir.
        /// Örn: fromCell="B2", toCell="M20"
        /// </summary>
        public static void AddChart(WorksheetPart worksheetPart, ChartDefinition def, string topLeftCell, string bottomRightCell)
        {
            // 1) Drawing part’i al/oluştur
            var drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();;

            // 2) Worksheet içine <drawing r:id="..."> ekle (yoksa)
            if (!worksheetPart.Worksheet.Elements<Drawing>().Any())
            {
                var drawing = new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) };
                // SheetData’dan sonra eklemek yaygın bir pratiktir
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                if (sheetData != null) sheetData.InsertAfterSelf(drawing);
                else worksheetPart.Worksheet.Append(drawing);

                worksheetPart.Worksheet.Save();
            }

            // 3) ChartPart oluştur
            var chartPart = drawingsPart.AddNewPart<ChartPart>();
            BuildChartContent(chartPart, def);

            // 4) Anchor (grafiği sheet’e yerleştir)
            drawingsPart.WorksheetDrawing ??= new WorksheetDrawing();
            UInt32Value nextNvId = GetNextDrawingId(drawingsPart.WorksheetDrawing);

            // Hücre referanslarını iki hücreli ankora çevir (A1, H20 gibi)
            var (fromCol, fromRow) = SplitCell(topLeftCell);
            var (toCol, toRow) = SplitCell(bottomRightCell);

            var twoCellAnchor = new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId(fromCol.ToString()),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId(fromRow.ToString()),
                    new Xdr.RowOffset("0")
                ),
                new Xdr.ToMarker(
                    new Xdr.ColumnId(toCol.ToString()),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId(toRow.ToString()),
                    new Xdr.RowOffset("0")
                ),
                new Xdr.GraphicFrame(
                    new Xdr.NonVisualGraphicFrameProperties(
                        new Xdr.NonVisualDrawingProperties { Id = nextNvId, Name = $"Chart {nextNvId}" },
                        new Xdr.NonVisualGraphicFrameDrawingProperties()
                    ),
                    new Xdr.Transform(
                        new A.Offset { X = 0, Y = 0 },
                        new A.Extents { Cx = 0, Cy = 0 }
                    ),
                    new A.Graphic(
                        new A.GraphicData(
                            new C.ChartReference { Id = drawingsPart.GetIdOfPart(chartPart) }
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
                    )
                ),
                new Xdr.ClientData()
            );


            drawingsPart.WorksheetDrawing.Append(twoCellAnchor);
            drawingsPart.WorksheetDrawing.Save();
        }

        private static UInt32Value GetNextDrawingId(Xdr.WorksheetDrawing drawing)
        {
            uint max = 1;
            foreach (var nv in drawing.Descendants<Xdr.NonVisualDrawingProperties>())
            {
                if (nv.Id != null && nv.Id.HasValue)
                    max = Math.Max(max, nv.Id.Value);
            }
            return max + 1;
        }


        private static (int col, int row) SplitCell(string cellRef)
        {
            // Örn "B12" -> (1,11) (0-based sütun, 0-based satır istiyorsanız ayarlayın)
            int i = 0;
            while (i < cellRef.Length && char.IsLetter(cellRef[i])) i++;
            var colLetters = cellRef.Substring(0, i).ToUpperInvariant();
            var rowDigits = cellRef.Substring(i);

            int col = 0;
            foreach (var ch in colLetters)
            {
                col = col * 26 + (ch - 'A' + 1);
            }
            col -= 1;                // 0-based’e çek
            int row = int.Parse(rowDigits) - 1; // 0-based’e çek
            return (col, row);
        }


        /// <summary>
        /// ChartPart içeriğini doldurur.
        /// </summary>
        public static void BuildChartContent(ChartPart chartPart, ChartDefinition def)
        {
            // ChartSpace kökü
            var chartSpace = new C.ChartSpace();

            // (İsteğe bağlı) dil/uyum bilgileri
            chartSpace.Append(new C.EditingLanguage { Val = "en-US" });

            // Chart gövdesi
            var chart = new C.Chart();

            // Başlık (BuildChartTitle: C.Title döndürmeli)
            if (def?.Title != null)
                chart.Title = BuildChartTitle(def.Title);

            // PlotArea (içi dolu olmalı)
            chart.PlotArea = BuildPlotArea(def);

            // Legend (BuildLegend: C.Legend döndürmeli)
            if (def?.Legend != null)
                chart.Append(BuildLegend(def));

            // Diğer yaygın bayraklar (opsiyonel)
            chart.PlotVisibleOnly = new C.PlotVisibleOnly { Val = def?.PlotVisibleOnly ?? true };
            if (def?.BlanksDisplayedAs != null)
            {
                chart.DisplayBlanksAs = new C.DisplayBlanksAs
                {
                    Val = def.BlanksDisplayedAs switch
                    {
                        BlanksDisplayedAs.Zero => C.DisplayBlanksAsValues.Zero,
                        BlanksDisplayedAs.Gap => C.DisplayBlanksAsValues.Gap,
                        _ => C.DisplayBlanksAsValues.Span
                    }
                };
            }

            // Chart nesnesini ChartSpace'e ekle
            chartSpace.Append(chart);
            // Kaydet
            chartPart.ChartSpace = chartSpace;
            chartPart.ChartSpace.Save();
        }

        // ---------------------------
        // Plot Area + Chart type
        // ---------------------------
        private static C.PlotArea BuildPlotArea(ChartDefinition def)
        {
            var plotArea = new C.PlotArea(
                new C.Layout()
            );

            switch (def.Type)
            {
                case ChartType.Column:
                    var col = new C.BarChart(
                        new C.BarDirection { Val = def.Type == ChartType.Bar ? C.BarDirectionValues.Bar : C.BarDirectionValues.Column },
                        new C.BarGrouping { Val = MapBarGrouping(def.GroupingType) },
                        new C.VaryColors { Val = def.VaryColors ?? false }
                    );

                    // Seriler
                    int colIdx = 0;
                    foreach (var s in def.Series)
                    {
                        col.Append(BuildBarOrLineSeries(s, colIdx++));
                    }

                    // Axis Ids
                    var axIdsC = EnsureAxes(plotArea, def, categorical: true);
                    col.Append(new C.AxisId() { Val = axIdsC.catId });
                    col.Append(new C.AxisId() { Val = axIdsC.valId });

                    plotArea.Append(col);
                    break;

                case ChartType.Line:
                    var line = new C.LineChart(
                        new C.Grouping { Val = MapLineGrouping(def.GroupingType) },
                        new C.VaryColors { Val = def.VaryColors ?? false }
                    );
                    int lineIdx = 0;
                    foreach (var s in def.Series)
                    {
                        line.Append(BuildBarOrLineSeries(s, lineIdx++, isLine: true));
                    }
                    var axIdsL = EnsureAxes(plotArea, def, categorical: true);
                    line.Append(new C.AxisId() { Val = axIdsL.catId });
                    line.Append(new C.AxisId() { Val = axIdsL.valId });

                    plotArea.Append(line);
                    break;

                case ChartType.Pie:
                    var pie = new C.PieChart(
                        new C.VaryColors() { Val = def.VaryColors ?? true }
                    );
                    int pieIdx = 0;
                    foreach (var s in def.Series)
                    {
                        pie.Append(BuildPieSeries(s, pieIdx++));
                    }
                    plotArea.Append(pie);
                    break;

                case ChartType.Scatter:
                    var scatter = new C.ScatterChart(
                        new C.ScatterStyle() { Val = MapScatterStyle(def.ScatterStyle) }
                    );
                    int scIdx = 0;
                    foreach (var s in def.Series)
                    {
                        scatter.Append(BuildScatterSeries(s, scIdx++));
                    }
                    var axIdsS = EnsureAxes(plotArea, def, categorical: false);
                    scatter.Append(new C.AxisId() { Val = axIdsS.catId }); // x
                    scatter.Append(new C.AxisId() { Val = axIdsS.valId }); // y

                    plotArea.Append(scatter);
                    break;

                default:
                    // Basit fallback: Column
                    var fallback = new C.BarChart(
                        new C.BarDirection() { Val = C.BarDirectionValues.Column },
                        new C.BarGrouping() { Val = C.BarGroupingValues.Clustered },
                        new C.VaryColors() { Val = def.VaryColors ?? false }
                    );
                    int fIdx = 0;
                    foreach (var s in def.Series)
                        fallback.Append(BuildBarOrLineSeries(s, fIdx++));

                    var axIdsF = EnsureAxes(plotArea, def, categorical: true);
                    fallback.Append(new C.AxisId() { Val = axIdsF.catId });
                    fallback.Append(new C.AxisId() { Val = axIdsF.valId });
                    plotArea.Append(fallback);
                    break;
            }

            return plotArea;
        }

        // ---------------------------
        // Series builders
        // ---------------------------
        private static OpenXmlElement BuildBarOrLineSeries(SeriesDefinition s, int idx, bool isLine = false)
        {
            var ser = isLine ? (OpenXmlCompositeElement)new C.LineChartSeries() : new C.BarChartSeries();
            ser.Append(
                new C.Index() { Val = (uint)idx },
                new C.Order() { Val = (uint)idx }
            );

            // Title
            if (!string.IsNullOrWhiteSpace(s?.Title?.Text))
                ser.Append(new C.SeriesText(new C.StringReference(new C.StringCache(new C.PointCount() { Val = 1U }, new C.StringPoint() { Index = 0U, NumericValue = new C.NumericValue(s.Title.Text) }))));

            // Category cache ve values cache
            var (catCache, valCache) = BuildCategoryAndValueCachesForCategorical(s);

            ser.Append(new C.CategoryAxisData(catCache));
            ser.Append(new C.Values(valCache));

            // Renk (basit: solid)
            var colorHex = s.Color ?? "#4472C4";
            ser.Append(BuildSolidColorShapeProps(colorHex));

            return ser;
        }

        private static OpenXmlElement BuildPieSeries(SeriesDefinition s, int idx)
        {
            var ser = new C.PieChartSeries(
                new C.Index() { Val = (uint)idx },
                new C.Order() { Val = (uint)idx }
            );

            if (!string.IsNullOrWhiteSpace(s?.Title?.Text))
                ser.Append(new C.SeriesText(new C.StringReference(new C.StringCache(new C.PointCount() { Val = 1U }, new C.StringPoint() { Index = 0U, NumericValue = new C.NumericValue(s.Title.Text) }))));

            var (catCache, valCache) = BuildCategoryAndValueCachesForCategorical(s);

            ser.Append(new C.CategoryAxisData(catCache));
            ser.Append(new C.Values(valCache));

            var colorHex = s.Color ?? "#4472C4";
            ser.Append(BuildSolidColorShapeProps(colorHex));

            return ser;
        }

        private static OpenXmlElement BuildScatterSeries(SeriesDefinition s, int idx)
        {
            var ser = new C.ScatterChartSeries(
                new C.Index() { Val = (uint)idx },
                new C.Order() { Val = (uint)idx }
            );

            if (!string.IsNullOrWhiteSpace(s?.Title?.Text))
                ser.Append(new C.SeriesText(new C.NumericValue(s.Title.Text)));

            // X ve Y cache
            var xs = s.Points?.Where(p => p.X.HasValue && p.Y.HasValue).ToList() ?? new List<PointDefinition>();
            var xCache = new C.NumberLiteral();
            var yCache = new C.NumberLiteral();

            xCache.Append(new C.FormatCode("General"));
            yCache.Append(new C.FormatCode("General"));

            uint n = 0;
            xCache.Append(new C.PointCount() { Val = (uint)xs.Count });
            yCache.Append(new C.PointCount() { Val = (uint)xs.Count });

            foreach (var p in xs)
            {
                xCache.Append(new C.NumericPoint() { Index = n, NumericValue = new C.NumericValue(p.X.Value.ToString(CultureInfo.InvariantCulture)) });
                yCache.Append(new C.NumericPoint() { Index = n, NumericValue = new C.NumericValue(p.Y.Value.ToString(CultureInfo.InvariantCulture)) });
                n++;
            }

            ser.Append(new C.XValues(xCache));
            ser.Append(new C.YValues(yCache));


            var colorHex = s.Color ?? "#4472C4";
            ser.Append(BuildSolidColorShapeProps(colorHex));

            return ser;
        }

        private static (C.StringLiteral, C.NumberLiteral)
        BuildCategoryAndValueCachesForCategorical(SeriesDefinition s)
        {
            var pts = s.Points ?? new List<PointDefinition>();

            var strLit = new C.StringLiteral();
            var numLit = new C.NumberLiteral();

            strLit.Append(new C.PointCount { Val = (uint)pts.Count });
            numLit.Append(new C.FormatCode("General"));
            numLit.Append(new C.PointCount { Val = (uint)pts.Count });

            for (uint i = 0; i < pts.Count; i++)
            {
                var cat = pts[(int)i].Category ?? string.Empty;
                var val = pts[(int)i].Value ?? 0d;

                strLit.Append(new C.StringPoint { Index = i, NumericValue = new C.NumericValue(cat) });
                numLit.Append(new C.NumericPoint { Index = i, NumericValue = new C.NumericValue(val.ToString(CultureInfo.InvariantCulture)) });
            }

            return (strLit, numLit);
        }


        private static OpenXmlElement BuildSolidColorShapeProps(string hex)
        {
            // Şekil/seri doldurma/çizgi rengi
            var rgb = NormalizeHex(hex);

            var spPr = new C.ChartShapeProperties(
                new A.SolidFill(new A.RgbColorModelHex() { Val = rgb }),
                new A.Outline(new A.SolidFill(new A.RgbColorModelHex() { Val = rgb }))
            );
            return spPr;
        }

        // ---------------------------
        // Axes (Category/Numeric)
        // ---------------------------
        private static (uint catId, uint valId) EnsureAxes(C.PlotArea plotArea, ChartDefinition def, bool categorical)
        {
            // Eksen Id’leri
            uint catId = 48650112U + (uint)new Random().Next(1, int.MaxValue / 4);
            uint valId = catId + 1;

            // Category Axis
            var catAx = new C.CategoryAxis(
                new C.AxisId() { Val = catId },
                new C.Scaling(new C.Orientation() { Val = C.OrientationValues.MinMax }),
                new C.Delete() { Val = (def.ShowCategoryAxis ? false : true) },
                new C.AxisPosition() { Val = C.AxisPositionValues.Bottom },
                new C.MajorTickMark() { Val = C.TickMarkValues.Outside },
                new C.MinorTickMark() { Val = C.TickMarkValues.None },
                new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo },
                new C.CrossingAxis() { Val = valId },
                new C.Crosses() { Val = C.CrossesValues.AutoZero },
                new C.AutoLabeled() { Val = true },
                new C.LabelAlignment() { Val = C.LabelAlignmentValues.Center },
                new C.LabelOffset() { Val = 100 }
            );

            // Value Axis
            var valAx = new C.ValueAxis(
                new C.AxisId() { Val = valId },
                new C.Scaling(new C.Orientation() { Val = C.OrientationValues.MinMax }),
                new C.Delete() { Val = (def.ShowValueAxis ? false : true) },
                new C.AxisPosition() { Val = C.AxisPositionValues.Left },
                new C.MajorGridlines(),
                new C.NumberingFormat() { FormatCode = "General", SourceLinked = true },
                new C.MajorTickMark() { Val = C.TickMarkValues.Outside },
                new C.MinorTickMark() { Val = C.TickMarkValues.None },
                new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo },
                new C.CrossingAxis() { Val = catId },
                new C.Crosses() { Val = C.CrossesValues.AutoZero },
                new C.CrossBetween() { Val = C.CrossBetweenValues.Between }
            );

            // Min/Max/MajorUnit uygula (ValueAxis)
            if (def.ValueAxis != null)
            {
                if (def.ValueAxis.Min.HasValue)
                    valAx.Scaling.Append(new C.MinAxisValue { Val = def.ValueAxis.Min.Value });

                if (def.ValueAxis.Max.HasValue)
                    valAx.Scaling.Append(new C.MaxAxisValue { Val = def.ValueAxis.Max.Value });


                if (def.ValueAxis.MajorUnit.HasValue)
                    valAx.Append(new C.MajorUnit() { Val = def.ValueAxis.MajorUnit.Value });
            }

            plotArea.Append(catAx);
            plotArea.Append(valAx);

            return (catId, valId);
        }

        // ---------------------------
        // Title / Legend
        // ---------------------------
        private static C.Title BuildChartTitle(TitleDefinition t)
        {
            if (t == null || string.IsNullOrWhiteSpace(t.Text))
                return new C.Title(new C.Overlay() { Val = false }); // boş başlık

            return new C.Title(
                new C.ChartText(
                    new C.RichText(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.Run(new A.Text(t.Text ?? string.Empty))
                        )
                    )
                ),
                new C.Overlay() { Val = false }
            );
        }

        private static C.Legend BuildLegend(ChartDefinition def)
        {
            var lg = new C.Legend(
                new C.LegendPosition() { Val = MapLegend(def.Legend?.Position ?? LegendPosition.Right) },
                new C.Overlay() { Val = false }
            );

            // Gizle/göster
            if (def.Legend != null && !def.Legend.Show)
                lg.Append(new C.Layout(new C.ManualLayout())); // gösterilmez gibi davranır (OpenXML'de direkt Delete yok)
            return lg;
        }

        // ---------------------------
        // Mappers
        // ---------------------------

        private static C.LegendPositionValues MapLegend(LegendPosition p) => p switch
        {
            LegendPosition.Top => C.LegendPositionValues.Top,
            LegendPosition.Bottom => C.LegendPositionValues.Bottom,
            LegendPosition.Left => C.LegendPositionValues.Left,
            LegendPosition.TopRight => C.LegendPositionValues.TopRight,
            _ => C.LegendPositionValues.Right
        };

        // Bar/Column için (ayrı enum!)
        private static C.BarGroupingValues MapBarGrouping(GroupingType? g) => g switch
        {
            GroupingType.Stacked => C.BarGroupingValues.Stacked,
            GroupingType.PercentStacked => C.BarGroupingValues.PercentStacked,
            _ => C.BarGroupingValues.Clustered
        };

        private static C.GroupingValues MapLineGrouping(GroupingType? g) => g switch
        {
            GroupingType.Stacked => C.GroupingValues.Stacked,
            GroupingType.PercentStacked => C.GroupingValues.PercentStacked,
            _ => C.GroupingValues.Standard
        };


        private static C.ScatterStyleValues MapScatterStyle(ScatterStyle s) => s switch
        {
            ScatterStyle.Line => C.ScatterStyleValues.Line,
            ScatterStyle.LineMarker => C.ScatterStyleValues.LineMarker,
            ScatterStyle.Smooth => C.ScatterStyleValues.Smooth,
            ScatterStyle.SmoothMarker => C.ScatterStyleValues.SmoothMarker,
            _ => C.ScatterStyleValues.Marker
        };

        // ---------------------------
        // Helpers
        // ---------------------------
        private static (int row, int col) ParseCellRef(string a1)
        {
            // Basit A1 -> (row,col) dönüşümü
            // Örn: B12 => col=2, row=12
            if (string.IsNullOrWhiteSpace(a1)) return (1, 1);

            string letters = new string(a1.ToUpper().Where(char.IsLetter).ToArray());
            string numbers = new string(a1.Where(char.IsDigit).ToArray());

            int col = 0;
            foreach (var ch in letters)
                col = col * 26 + (ch - 'A' + 1);

            int row = 1;
            int.TryParse(numbers, out row);
            if (row <= 0) row = 1;
            if (col <= 0) col = 1;

            return (row, col);
        }


        private static string NormalizeHex(string hex)
        {
            if (string.IsNullOrWhiteSpace(hex)) return "4472C4";
            hex = hex.Trim();
            if (hex.StartsWith("#")) hex = hex.Substring(1);
            if (hex.Length == 3)
            {
                // #RGB → #RRGGBB
                hex = new string(new[] { hex[0], hex[0], hex[1], hex[1], hex[2], hex[2] });
            }
            if (hex.Length != 6) return "4472C4";
            return hex.ToUpperInvariant();
        }
    }


}

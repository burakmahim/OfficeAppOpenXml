using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using System.Globalization;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using System.Xml;
using System.Text.RegularExpressions;

namespace OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents
{
    public class ChartComponent
    {
        class ChartDefinition
        {
            public bool IsAreaBased
            {
                get
                {
                    switch (Type)
                    {
                        case ChartType.Area:
                        case ChartType.Bar:
                        case ChartType.Column:
                        case ChartType.Pie:
                        case ChartType.Doughnut:
                        case ChartType.Bubble:
                            return true;
                        case ChartType.Radar:
                            return RadarStyle == RadarStyle.Filled;
                        default:
                            return false;
                    }
                }
            }
            public bool IsLineBased
            {
                get
                {
                    switch (Type)
                    {
                        case ChartType.Line:
                            return true;
                        case ChartType.Scatter:
                            return ScatterStyle is ScatterStyle.Line
                                                or ScatterStyle.LineMarker
                                                or ScatterStyle.Smooth
                                                or ScatterStyle.SmoothMarker;
                        case ChartType.Radar:
                            return RadarStyle is RadarStyle.Standard or RadarStyle.Marker;
                        default:
                            return false;
                    }
                }
            }
            public ChartType Type { get; set; }
            public TitleDefinition Title { get; set; }
            public Dimension Width { get; set; }
            public Dimension Height { get; set; }
            public int? GapWidth { get; set; }   // Bar, Column only
            public int? Overlap { get; set; }   // Bar, Column only
            public ScatterStyle ScatterStyle { get; set; }
            public RadarStyle RadarStyle { get; set; }
            public MarkerDefinition Marker { get; set; }   // Line, Scatter, Radar only
            public List<SeriesDefinition> Series { get; set; } = [];
            public AxisDefinition CategoryAxis { get; set; }
            public AxisDefinition ValueAxis { get; set; }
            public LegendDefinition Legend { get; set; }
            public GridDefinition Grid { get; set; }
            public ValueLabelDefinition ValueLabels { get; set; }
            public ChartLayoutDefinition Layout { get; set; }
            public ThreeDViewDefinition ThreeDView { get; set; }
            public DataTableDefinition DataTable { get; set; }
            public FormatDefinition PlotAreaFormat { get; set; }
            public FormatDefinition ChartAreaFormat { get; set; }
            public TextFormatDefinition TextFormat { get; set; }
            public FormatDefinition SeriesDefaultFormat { get; set; }

            public bool? AutoTitle { get; set; } //default: false
            public BlanksDisplayedAs? BlanksDisplayedAs { get; set; }
            public bool? PlotVisibleOnly { get; set; }
            public bool? RoundedCorners { get; set; }
            public GroupingType? GroupingType { get; set; }
            public bool? VaryColors { get; set; }
            public bool? ShowDataLabelsOverMaximum { get; set; }
            public bool CreateNewParagraph { get; set; } //default true
            public bool ShowCategoryAxis { get; set; } //default true
            public bool ShowValueAxis { get; set; } //default true
            public int? FirstSliceAngle { get; set; } // Pie/Doughnut: 0..360
            public int? DoughnutHoleSize { get; set; } // Doughnut: 10..90
            public int? BubbleScale { get; set; } // Bubble: 0..300 (yüzde)
            public bool? Bubble3D { get; set; } // Bubble: true/false
        }

        private static ChartDefinition GetChartDefinitionFromXml(XElement chartNode)
        {
            ChartDefinition chartDefinition = new ChartDefinition();

            //Chart attributes----------------------------------------------------------------------

            Nullable<ChartType> chartType = XElementAttributeGetter.AsEnum<ChartType>(chartNode, "type");
            if (chartType.HasValue)
                chartDefinition.Type = chartType.Value;

            Nullable<RadarStyle> radarStyle = XElementAttributeGetter.AsEnum<RadarStyle>(chartNode, "radar-style");
            if (radarStyle.HasValue)
                chartDefinition.RadarStyle = radarStyle.Value;

            Nullable<GroupingType> groupingType = XElementAttributeGetter.AsEnum<GroupingType>(chartNode, "grouping-type");
            if (groupingType.HasValue)
                chartDefinition.GroupingType = groupingType.Value;

            Nullable<ScatterStyle> scatterStyle = XElementAttributeGetter.AsEnum<ScatterStyle>(chartNode, "scatter-style");
            if (scatterStyle.HasValue)
                chartDefinition.ScatterStyle = scatterStyle.Value;

            Nullable<BlanksDisplayedAs> displayBlanksAs = XElementAttributeGetter.AsEnum<BlanksDisplayedAs>(chartNode, "display-blank-as");
            if (displayBlanksAs.HasValue)
                chartDefinition.BlanksDisplayedAs = displayBlanksAs.Value;

            XElementAttributeGetter.AsBool(chartNode, "vary-colors", out bool varyColorsValue);
            chartDefinition.VaryColors = varyColorsValue;

            XElementAttributeGetter.AsBool(chartNode, "auto-title", out bool autoTitleValue, defaultValue: false);
            chartDefinition.AutoTitle = autoTitleValue;

            XElementAttributeGetter.AsBool(chartNode, "show-category-axis", out bool showCategoryAxisValue, defaultValue: true);
            chartDefinition.ShowCategoryAxis = showCategoryAxisValue;

            XElementAttributeGetter.AsBool(chartNode, "show-value-axis", out bool showValueAxisValue, defaultValue: true);
            chartDefinition.ShowValueAxis = showValueAxisValue;

            XElementAttributeGetter.AsBool(chartNode, "plot-visible-only", out bool plotVisibleOnlyValue, defaultValue: false);
            chartDefinition.PlotVisibleOnly = plotVisibleOnlyValue;

            XElementAttributeGetter.AsBool(chartNode, "rounded-corners", out bool roundedCornersValue, defaultValue: false);
            chartDefinition.RoundedCorners = roundedCornersValue;

            XElementAttributeGetter.AsBool(chartNode, "show-data-labels-over-maximum", out bool showDataLabelsOverMaximumValue, defaultValue: false);
            chartDefinition.ShowDataLabelsOverMaximum = showDataLabelsOverMaximumValue;

            XElementAttributeGetter.AsBool(chartNode, "bubble-3d", out bool bubble3DValue, defaultValue: false);
            chartDefinition.Bubble3D = bubble3DValue;

            if (XElementAttributeGetter.AsInt32(chartNode, "overlap", out int overlapValue))
                chartDefinition.Overlap = overlapValue;

            if (XElementAttributeGetter.AsInt32(chartNode, "gap-width", out int gapWidthValue))
                chartDefinition.GapWidth = gapWidthValue;

            if (XElementAttributeGetter.AsInt32(chartNode, "first-slice-angle", out int firstSliceAngleValue))
                chartDefinition.FirstSliceAngle = firstSliceAngleValue;

            if (XElementAttributeGetter.AsInt32(chartNode, "doughnut-hole-size", out int doughnutHoleSizeValue))
                chartDefinition.DoughnutHoleSize = doughnutHoleSizeValue;

            if (XElementAttributeGetter.AsInt32(chartNode, "bubble-scale", out int bubbleScaleValue))
                chartDefinition.BubbleScale = bubbleScaleValue;

            //category-axis 

            XElement? categoryAxisNode = chartNode.Element("category-axis");
            if (categoryAxisNode != null)
            {
                chartDefinition.CategoryAxis = new AxisDefinition();

                if (XElementAttributeGetter.AsDouble(categoryAxisNode, "min", out double minValue))
                    chartDefinition.CategoryAxis.Min = minValue;

                if (XElementAttributeGetter.AsDouble(categoryAxisNode, "max", out double maxValue))
                    chartDefinition.CategoryAxis.Max = maxValue;

                if (XElementAttributeGetter.AsDouble(categoryAxisNode, "major-unit", out double majorUnit))
                    chartDefinition.CategoryAxis.MajorUnit = majorUnit;

                XElement? formatNode = categoryAxisNode.Element("format");
                if (formatNode != null)
                {
                    chartDefinition.CategoryAxis.AxisLineFormat = new FormatDefinition();

                    XElement? fillNode = formatNode.Element("fill");
                    if (fillNode != null)
                    {
                        chartDefinition.CategoryAxis.AxisLineFormat.FillDefinition = new FillDefinition();

                        XElement? solidNode = fillNode.Element("solid");
                        if (solidNode != null)
                        {
                            chartDefinition.CategoryAxis.AxisLineFormat.FillDefinition.SolidFillDefinition = new SolidDefinition();

                            ColorA color = ColorA.Parse(solidNode, "color", "transparency");
                            if (color != null)
                                chartDefinition.CategoryAxis.AxisLineFormat.FillDefinition.SolidFillDefinition.Color = color;
                        }

                        XElement? patternNode = fillNode.Element("pattern");
                        if (patternNode != null)
                        {
                            chartDefinition.CategoryAxis.AxisLineFormat.FillDefinition.PatternFillDefinition = new PatternDefinition();

                            Nullable<PatternFillPreset> patternType = XElementAttributeGetter.AsEnum<PatternFillPreset>(patternNode, "preset");
                            if (patternType.HasValue)
                                chartDefinition.CategoryAxis.AxisLineFormat.FillDefinition.PatternFillDefinition.Preset = patternType.Value;

                            ColorA foregroundColor = ColorA.Parse(patternNode, "foreground-color", "foreground-transparency");
                            if (foregroundColor != null)
                                chartDefinition.CategoryAxis.AxisLineFormat.FillDefinition.PatternFillDefinition.ForegroundColor = foregroundColor;

                            ColorA backgroundColor = ColorA.Parse(patternNode, "background-color", "background-transparency");
                            if (backgroundColor != null)
                                chartDefinition.CategoryAxis.AxisLineFormat.FillDefinition.PatternFillDefinition.BackgroundColor = backgroundColor;
                        }

                        XElement? gradientNode = fillNode.Element("gradient");
                        if (gradientNode != null)
                        {
                            GradientDefinition gradientDef = new GradientDefinition();

                            if (XElementAttributeGetter.AsInt32(gradientNode, "angle", out int angleValue))
                                gradientDef.Angle = angleValue;

                            XElementAttributeGetter.AsBool(gradientNode, "scaled", out bool scaledValue);
                            gradientDef.Scaled = scaledValue;

                            foreach (XElement stopNode in gradientNode.Elements("stop"))
                            {
                                var gradientStop = new GradientStopDefinition();

                                if (XElementAttributeGetter.AsDouble(stopNode, "position", out double positionValue))
                                    gradientStop.Position = (int)Math.Round(positionValue * 1000);
                                else
                                    gradientStop.Position = 0; // default

                                ColorA? color = ColorA.Parse(stopNode, "color", "transparency");
                                if (color != null)
                                    gradientStop.Color = color;

                                if (gradientStop.Color != null)
                                    gradientDef.Stops.Add(gradientStop);
                            }

                            chartDefinition.CategoryAxis.AxisLineFormat.FillDefinition.GradientFillDefinition = gradientDef;
                        }

                    }

                    XElement? lineNode = formatNode.Element("line");
                    if (lineNode != null)
                    {
                        chartDefinition.CategoryAxis.AxisLineFormat.LineDefinition = new LineDefinition();

                        XElementAttributeGetter.AsBool(lineNode, "visible", out bool visibleValue);
                        chartDefinition.CategoryAxis.AxisLineFormat.LineDefinition.Visible = visibleValue;

                        Nullable<DashPreset> dashPreset = XElementAttributeGetter.AsEnum<DashPreset>(lineNode, "dash");
                        if (dashPreset.HasValue)
                            chartDefinition.CategoryAxis.AxisLineFormat.LineDefinition.DashPreset = dashPreset.Value;

                        Nullable<CompoundLinePreset> compoundLine = XElementAttributeGetter.AsEnum<CompoundLinePreset>(lineNode, "compound");
                        if (compoundLine.HasValue)
                            chartDefinition.CategoryAxis.AxisLineFormat.LineDefinition.CompoundPreset = compoundLine.Value;

                        Nullable<LineCapPreset> lineCap = XElementAttributeGetter.AsEnum<LineCapPreset>(lineNode, "cap");
                        if (lineCap.HasValue)
                            chartDefinition.CategoryAxis.AxisLineFormat.LineDefinition.CapPreset = lineCap.Value;

                        Nullable<LineJoinPreset> join = XElementAttributeGetter.AsEnum<LineJoinPreset>(lineNode, "join");
                        if (join.HasValue)
                            chartDefinition.CategoryAxis.AxisLineFormat.LineDefinition.JoinPreset = join.Value;

                        Nullable<LineEndPreset> beginArrowType = XElementAttributeGetter.AsEnum<LineEndPreset>(lineNode, "begin-arrow-type");
                        if (beginArrowType.HasValue)
                            chartDefinition.CategoryAxis.AxisLineFormat.LineDefinition.Begin_Arrow_Type = beginArrowType.Value;

                        Nullable<LineEndWidthPreset> beginArrowWidth = XElementAttributeGetter.AsEnum<LineEndWidthPreset>(lineNode, "begin-arrow-width");
                        if (beginArrowWidth.HasValue)
                            chartDefinition.CategoryAxis.AxisLineFormat.LineDefinition.Begin_Arrow_Width = beginArrowWidth.Value;

                        Nullable<LineEndLengthPreset> beginArrowLength = XElementAttributeGetter.AsEnum<LineEndLengthPreset>(lineNode, "begin-arrow-length");
                        if (beginArrowLength.HasValue)
                            chartDefinition.CategoryAxis.AxisLineFormat.LineDefinition.Begin_Arrow_Length = beginArrowLength.Value;

                        Nullable<LineEndPreset> endArrowType = XElementAttributeGetter.AsEnum<LineEndPreset>(lineNode, "end-arrow-type");
                        if (endArrowType.HasValue)
                            chartDefinition.CategoryAxis.AxisLineFormat.LineDefinition.End_Arrow_Type = endArrowType.Value;

                        Nullable<LineEndWidthPreset> endArrowWidth = XElementAttributeGetter.AsEnum<LineEndWidthPreset>(lineNode, "end-arrow-width");
                        if (endArrowWidth.HasValue)
                            chartDefinition.CategoryAxis.AxisLineFormat.LineDefinition.End_Arrow_Width = endArrowWidth.Value;

                        Nullable<LineEndLengthPreset> endArrowLength = XElementAttributeGetter.AsEnum<LineEndLengthPreset>(lineNode, "end-arrow-length");
                        if (endArrowLength.HasValue)
                            chartDefinition.CategoryAxis.AxisLineFormat.LineDefinition.End_Arrow_Length = endArrowLength.Value;

                        if (XElementAttributeGetter.AsInt32(lineNode, "join-miter-limit", out int joinMiterLimitValue))
                        {
                            int clampedValue = clamp(joinMiterLimitValue, 1, 500);
                            chartDefinition.CategoryAxis.AxisLineFormat.LineDefinition.JoinMiterLimit = clampedValue;
                        }

                    }
                }

            }
            return chartDefinition;
        }

        private static C.Chart GetChartCommon(ChartDefinition chartDefinition, XElement chartNode, C.PlotArea plotArea, bool allowView3D = true)
        {
            C.Chart chart = new C.Chart();

            string chartTitle = chartNode.Attribute("title")?.Value ?? "";

            A.Text text = new A.Text(chartTitle);
            A.Run run = new A.Run(text);
            A.Paragraph paragraph = new A.Paragraph(run);
            C.RichText richText = new C.RichText(new A.BodyProperties(), new A.ListStyle(), paragraph);
            C.ChartText chartText = new C.ChartText(richText);
            C.Title title = new C.Title(chartText, new Overlay() { Val = false });

            chart.Append(title);
            chart.Append(plotArea);
            chart.Append(new DisplayBlanksAs() { Val = mapBlanksDisplayedAs(chartDefinition.BlanksDisplayedAs ?? BlanksDisplayedAs.Gap) });
            chart.Append(new AutoTitleDeleted() { Val = !chartDefinition.AutoTitle ?? false });
            chart.Append(new PlotVisibleOnly() { Val = chartDefinition.PlotVisibleOnly ?? false });
            chart.Append(new ShowDataLabelsOverMaximum() { Val = chartDefinition.ShowDataLabelsOverMaximum ?? true });

            return chart;
        }

        public static class XElementAttributeGetter
        {
            public static string AsString(XElement element, string attributeName)
            {
                return element?.Attribute(attributeName)?.Value ?? string.Empty;
            }

            public static bool AsDouble(XElement element, string attributeName, out double result)
            {
                result = 0;
                string value = AsString(element, attributeName);
                return double.TryParse(value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out result);
            }
            public static bool AsBool(XElement element, string attributeName, out bool result, bool defaultValue = false)
            {
                string value = AsString(element, attributeName);
                if (bool.TryParse(value, out result))
                    return true;
                result = defaultValue;
                return false;
            }
            public static T? AsEnum<T>(XElement element, string attributeName) where T : struct, Enum
            {
                string value = AsString(element, attributeName);
                if (string.IsNullOrWhiteSpace(value))
                    return null;

                if (Enum.TryParse<T>(value, true, out T result))
                    return result;

                return null;
            }
            public static bool AsInt32(XElement element, string attributeName, out int result)
            {
                result = 0;
                string value = AsString(element, attributeName);
                return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out result);
            }
        }
        public static void AddChart(WorksheetPart worksheetPart, XElement chartNode)
        {
            ChartDefinition chartDefinition = GetChartDefinitionFromXml(chartNode);

            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            if (drawingsPart.WorksheetDrawing == null)
                drawingsPart.WorksheetDrawing = new WorksheetDrawing();

            ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
            string relId = drawingsPart.GetIdOfPart(chartPart);

            C.ChartSpace chartSpace = new C.ChartSpace();
            chartSpace.Append(new EditingLanguage() { Val = "tr-TR" });

            C.Chart? chart = null;

            switch (chartDefinition.Type)
            {
                case ChartType.Bar:
                    chart = getBarChart(chartDefinition, chartNode, BarDirectionValues.Bar);
                    break;
                case ChartType.Column:
                    chart = getBarChart(chartDefinition, chartNode, BarDirectionValues.Column);
                    break;
                case ChartType.Line:
                    chart = getLineChart(chartDefinition, chartNode);
                    break;
                case ChartType.Area:
                    chart = getAreaChart(chartDefinition, chartNode);
                    break;
                case ChartType.Pie:
                    chart = getPieChart(chartDefinition, chartNode);
                    break;
                case ChartType.Doughnut:
                    chart = getDoughnutChart(chartDefinition, chartNode);
                    break;
                case ChartType.Scatter:
                    chart = getScatterChart(chartDefinition, chartNode);
                    break;
                case ChartType.Bubble:
                    chart = getBubbleChart(chartDefinition, chartNode);
                    break;
                case ChartType.Radar:
                    chart = getRadarChart(chartDefinition, chartNode);
                    break;
                default:
                    break;
            }

            if (chart is null)
                return;

            chartSpace.Append(new RoundedCorners() { Val = chartDefinition.RoundedCorners ?? false });

            chartSpace.Append(chart);

            chartPart.ChartSpace = chartSpace;

            chartPart.ChartSpace.Save();

            TwoCellAnchor twoCellAnchor = new TwoCellAnchor();
            twoCellAnchor.Append(new Xdr.FromMarker(
                new Xdr.ColumnId("1"), new Xdr.ColumnOffset("0"),
                new Xdr.RowId("1"), new Xdr.RowOffset("0")
            ));

            twoCellAnchor.Append(new Xdr.ToMarker(
                new Xdr.ColumnId("8"), new Xdr.ColumnOffset("0"),
                new Xdr.RowId("15"), new Xdr.RowOffset("0")
            ));

            GraphicFrame graphicFrame = new GraphicFrame();
            graphicFrame.Append(new NonVisualGraphicFrameProperties(
                new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "Chart" + Guid.NewGuid() },
                new NonVisualGraphicFrameDrawingProperties()
            ));

            graphicFrame.Append(new Transform(
                new A.Offset() { X = 0, Y = 0 },
                new A.Extents() { Cx = 0, Cy = 0 }
            ));

            graphicFrame.Append(new A.Graphic(
                new A.GraphicData(
                    new C.ChartReference() { Id = relId }
                )
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
            ));

            graphicFrame.Append(new ClientData());
            twoCellAnchor.Append(graphicFrame);
            drawingsPart.WorksheetDrawing.Append(twoCellAnchor);
            drawingsPart.WorksheetDrawing.Save();

            if (!worksheetPart.Worksheet.Elements<Drawing>().Any())
            {
                worksheetPart.Worksheet.Append(new Drawing() { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            worksheetPart.Worksheet.Save();
        }

        private static C.Chart getBarChart(ChartDefinition chartDefinition, XElement chartNode, BarDirectionValues direction)
        {
            uint categoryAxisId = getSafeId();
            uint valueAxisId = getSafeId();

            C.PlotArea plotArea = new C.PlotArea();
            plotArea.Append(new Layout());

            BarChart barChart = new BarChart()
            {
                BarDirection = new BarDirection() { Val = direction },
                BarGrouping = new BarGrouping() { Val = mapBarGrouping(chartDefinition.GroupingType ?? GroupingType.Standard) },
                VaryColors = new VaryColors() { Val = chartDefinition.VaryColors ?? false }
            };

            addAxisIds(barChart, categoryAxisId, valueAxisId);

            uint seriesIndex = 0;

            foreach (XElement seriesNode in chartNode.Elements("series"))
            {
                BarChartSeries series = new BarChartSeries(
                    new C.Index() { Val = seriesIndex },
                    new C.Order() { Val = seriesIndex },
                    new C.SeriesText(new C.NumericValue(seriesNode.Attribute("name")?.Value ?? ""))
                );

                List<XElement> points = seriesNode.Elements("point").ToList();
                uint pointCount = (uint)points.Count;

                CategoryAxisData catAxisData = new CategoryAxisData();
                C.Values values = new C.Values();

                StringLiteral stringLiteral = new StringLiteral();
                NumberLiteral numberLiteral = new NumberLiteral();

                stringLiteral.Append(new PointCount() { Val = pointCount });
                numberLiteral.Append(new PointCount() { Val = pointCount });

                for (int i = 0; i < pointCount; i++)
                {
                    string categoryName = points[i].Attribute("category")?.Value ?? $"Kategori {i + 1}";
                    string valueStr = points[i].Attribute("value")?.Value ?? "0";

                    stringLiteral.Append(new StringPoint() { Index = (uint)i, NumericValue = new C.NumericValue(categoryName) });
                    numberLiteral.Append(new NumericPoint() { Index = (uint)i, NumericValue = new C.NumericValue(valueStr) });
                }

                catAxisData.Append(stringLiteral);
                values.Append(numberLiteral);

                series.Append(catAxisData);
                series.Append(values);

                barChart.Append(series);
                seriesIndex++;
            }

            if (chartDefinition.Overlap.HasValue)
                barChart.Append(new Overlap() { Val = new SByteValue((sbyte)chartDefinition.Overlap) });

            if (chartDefinition.GapWidth.HasValue)
                barChart.Append(new GapWidth() { Val = new UInt16Value((ushort)chartDefinition.GapWidth) });

            plotArea.Append(barChart);

            AxisPositionValues categoryAxisPosition = direction == BarDirectionValues.Bar ? AxisPositionValues.Left : AxisPositionValues.Bottom;
            AxisPositionValues valueAxisPosition = direction == BarDirectionValues.Bar ? AxisPositionValues.Bottom : AxisPositionValues.Left;

            addCategoryAxis(chartDefinition, chartNode, plotArea, categoryAxisId, valueAxisId, categoryAxisPosition);
            addValueAxis(chartDefinition, chartNode, plotArea, categoryAxisId, valueAxisId, valueAxisPosition);

            return GetChartCommon(chartDefinition, chartNode, plotArea);
        }
        private static C.Chart getLineChart(ChartDefinition chartDefinition, XElement chartNode)
        {
            uint categoryAxisId = getSafeId();
            uint valueAxisId = getSafeId();

            C.PlotArea plotArea = new C.PlotArea();
            plotArea.Append(new Layout());

            LineChart lineChart = new LineChart()
            {
                Grouping = new Grouping() { Val = GroupingValues.Standard },
                VaryColors = new VaryColors() { Val = chartDefinition.VaryColors ?? false }
            };

            addAxisIds(lineChart, categoryAxisId, valueAxisId);

            uint seriesIndex = 0;

            foreach (XElement seriesNode in chartNode.Elements("series"))
            {
                LineChartSeries series = new LineChartSeries(
                    new C.Index() { Val = seriesIndex },
                    new Order() { Val = seriesIndex },
                    new C.SeriesText(new C.NumericValue(seriesNode.Attribute("name")?.Value ?? ""))
                );

                List<XElement> points = seriesNode.Elements("point").ToList();
                uint pointCount = (uint)points.Count;

                CategoryAxisData catAxisData = new CategoryAxisData();
                C.Values values = new C.Values();

                StringLiteral stringLiteral = new StringLiteral();
                NumberLiteral numberLiteral = new NumberLiteral();

                stringLiteral.Append(new PointCount() { Val = pointCount });
                numberLiteral.Append(new PointCount() { Val = pointCount });

                for (int i = 0; i < pointCount; i++)
                {
                    string categoryName = points[i].Attribute("category")?.Value ?? $"Kategori {i + 1}";
                    string valueStr = points[i].Attribute("value")?.Value ?? "0";

                    stringLiteral.Append(new StringPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new C.NumericValue(categoryName)
                    });

                    numberLiteral.Append(new NumericPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new C.NumericValue(valueStr)
                    });
                }

                catAxisData.Append(stringLiteral);
                values.Append(numberLiteral);

                series.Append(catAxisData);
                series.Append(values);

                lineChart.Append(series);
                seriesIndex++;
            }

            plotArea.Append(lineChart);

            addCategoryAxis(chartDefinition, chartNode, plotArea, categoryAxisId, valueAxisId);
            addValueAxis(chartDefinition, chartNode, plotArea, valueAxisId, categoryAxisId);

            return GetChartCommon(chartDefinition, chartNode, plotArea);
        }
        private static C.Chart getAreaChart(ChartDefinition chartDefinition, XElement chartNode)
        {
            uint categoryAxisId = getSafeId();
            uint valueAxisId = getSafeId();

            C.PlotArea plotArea = new C.PlotArea();
            plotArea.Append(new Layout());

            AreaChart areaChart = new()
            {
                Grouping = new Grouping() { Val = GroupingValues.Standard },
                VaryColors = new VaryColors() { Val = chartDefinition.VaryColors ?? false }
            };

            addAxisIds(areaChart, categoryAxisId, valueAxisId);

            uint seriesIndex = 0;

            foreach (XElement seriesNode in chartNode.Elements("series"))
            {
                AreaChartSeries series = new AreaChartSeries(
                    new C.Index() { Val = seriesIndex },
                    new Order() { Val = seriesIndex },
                    new C.SeriesText(new C.NumericValue(seriesNode.Attribute("name")?.Value ?? ""))
                );

                List<XElement> points = seriesNode.Elements("point").ToList();
                uint pointCount = (uint)points.Count;

                CategoryAxisData catAxisData = new CategoryAxisData();
                C.Values values = new C.Values();

                StringLiteral stringLiteral = new StringLiteral();
                NumberLiteral numberLiteral = new NumberLiteral();

                stringLiteral.Append(new PointCount() { Val = pointCount });
                numberLiteral.Append(new PointCount() { Val = pointCount });

                for (int i = 0; i < pointCount; i++)
                {
                    string categoryName = points[i].Attribute("category")?.Value ?? $"Kategori {i + 1}";
                    string valueStr = points[i].Attribute("value")?.Value ?? "0";

                    stringLiteral.Append(new StringPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new C.NumericValue(categoryName)
                    });

                    numberLiteral.Append(new NumericPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new C.NumericValue(valueStr)
                    });
                }

                catAxisData.Append(stringLiteral);
                values.Append(numberLiteral);

                series.Append(catAxisData);
                series.Append(values);

                areaChart.Append(series);
                seriesIndex++;
            }

            plotArea.Append(areaChart);

            addCategoryAxis(chartDefinition, chartNode, plotArea, categoryAxisId, valueAxisId);
            addValueAxis(chartDefinition, chartNode, plotArea, categoryAxisId, valueAxisId);

            return GetChartCommon(chartDefinition, chartNode, plotArea);
        }
        private static C.Chart getPieChart(ChartDefinition chartDefinition, XElement chartNode)
        {

            C.PlotArea plotArea = new C.PlotArea();
            plotArea.Append(new Layout());

            PieChart pieChart = new(new VaryColors() { Val = chartDefinition.VaryColors ?? false });

            uint seriesIndex = 0;

            foreach (XElement seriesNode in chartNode.Elements("series"))
            {
                PieChartSeries series = new PieChartSeries(
                    new C.Index() { Val = seriesIndex },
                    new Order() { Val = seriesIndex },
                    new C.SeriesText(new C.NumericValue(seriesNode.Attribute("name")?.Value ?? ""))
                );

                List<XElement> points = seriesNode.Elements("point").ToList();
                uint pointCount = (uint)points.Count;

                CategoryAxisData catAxisData = new CategoryAxisData();
                C.Values values = new C.Values();

                StringLiteral stringLiteral = new StringLiteral();
                NumberLiteral numberLiteral = new NumberLiteral();

                stringLiteral.Append(new PointCount() { Val = pointCount });
                numberLiteral.Append(new PointCount() { Val = pointCount });

                for (int i = 0; i < pointCount; i++)
                {
                    string categoryName = points[i].Attribute("category")?.Value ?? $"Kategori {i + 1}";
                    string valueStr = points[i].Attribute("value")?.Value ?? "0";

                    stringLiteral.Append(new StringPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new C.NumericValue(categoryName)
                    });

                    numberLiteral.Append(new NumericPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new C.NumericValue(valueStr)
                    });
                }

                catAxisData.Append(stringLiteral);
                values.Append(numberLiteral);

                series.Append(catAxisData);
                series.Append(values);

                pieChart.Append(series);
                seriesIndex++;
            }

            if (chartDefinition.FirstSliceAngle.HasValue)
                pieChart.Append(new C.FirstSliceAngle() { Val = (UInt16Value)(ushort)clamp(chartDefinition.FirstSliceAngle.Value, 0, 360) });

            plotArea.Append(pieChart);

            return GetChartCommon(chartDefinition, chartNode, plotArea);
        }
        private static C.Chart getDoughnutChart(ChartDefinition chartDefinition, XElement chartNode)
        {

            C.PlotArea plotArea = new C.PlotArea();
            plotArea.Append(new Layout());

            DoughnutChart doughnutChart = new(new VaryColors() { Val = chartDefinition.VaryColors ?? false });

            uint seriesIndex = 0;

            foreach (XElement seriesNode in chartNode.Elements("series"))
            {
                PieChartSeries series = new PieChartSeries(
                    new C.Index() { Val = seriesIndex },
                    new Order() { Val = seriesIndex },
                    new C.SeriesText(new C.NumericValue(seriesNode.Attribute("name")?.Value ?? ""))
                );

                List<XElement> points = seriesNode.Elements("point").ToList();
                uint pointCount = (uint)points.Count;

                CategoryAxisData catAxisData = new CategoryAxisData();
                C.Values values = new C.Values();

                StringLiteral stringLiteral = new StringLiteral();
                NumberLiteral numberLiteral = new NumberLiteral();

                stringLiteral.Append(new PointCount() { Val = pointCount });
                numberLiteral.Append(new PointCount() { Val = pointCount });

                for (int i = 0; i < pointCount; i++)
                {
                    string categoryName = points[i].Attribute("category")?.Value ?? $"Kategori {i + 1}";
                    string valueStr = points[i].Attribute("value")?.Value ?? "0";

                    stringLiteral.Append(new StringPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new C.NumericValue(categoryName)
                    });

                    numberLiteral.Append(new NumericPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new C.NumericValue(valueStr)
                    });
                }

                catAxisData.Append(stringLiteral);
                values.Append(numberLiteral);

                series.Append(catAxisData);
                series.Append(values);

                doughnutChart.Append(series);
                seriesIndex++;
            }

            if (chartDefinition.FirstSliceAngle.HasValue)
                doughnutChart.Append(new C.FirstSliceAngle() { Val = (UInt16Value)(ushort)clamp(chartDefinition.FirstSliceAngle.Value, 0, 360) });

            doughnutChart.Append(new C.HoleSize() { Val = (ByteValue)(byte)clamp(chartDefinition.DoughnutHoleSize ?? 50, 10, 90) });

            plotArea.Append(doughnutChart);

            return GetChartCommon(chartDefinition, chartNode, plotArea);
        }
        private static C.Chart getScatterChart(ChartDefinition chartDefinition, XElement chartNode)
        {
            uint categoryAxisId = getSafeId();
            uint valueAxisId = getSafeId();

            C.PlotArea plotArea = new C.PlotArea();
            plotArea.Append(new Layout());

            ScatterChart scatterChart = new(
                new C.ScatterStyle() { Val = mapScatterStyle(chartDefinition.ScatterStyle) },
                new C.VaryColors() { Val = chartDefinition.VaryColors ?? false }
            );

            uint seriesIndex = 0;

            foreach (XElement seriesNode in chartNode.Elements("series"))
            {
                ScatterChartSeries series = new ScatterChartSeries(
                    new C.Index() { Val = seriesIndex },
                    new Order() { Val = seriesIndex },
                    new C.SeriesText(new C.NumericValue(seriesNode.Attribute("name")?.Value ?? ""))
                );

                List<XElement> points = seriesNode.Elements("point").ToList();
                uint pointCount = (uint)points.Count;

                NumberLiteral numberLiteralX = new NumberLiteral();
                numberLiteralX.Append(new C.PointCount() { Val = (uint)pointCount });
                for (int i = 0; i < pointCount; i++)
                {
                    string? x = points[i].Attribute("x")?.Value;

                    numberLiteralX.Append(new C.NumericPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new C.NumericValue(x)
                    });
                }

                NumberLiteral numberLiteralY = new NumberLiteral();
                numberLiteralY.Append(new C.PointCount() { Val = (uint)pointCount });
                for (int i = 0; i < pointCount; i++)
                {
                    string? y = points[i].Attribute("y")?.Value;

                    numberLiteralY.Append(new C.NumericPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new C.NumericValue(y)
                    });
                }

                series.Append(new C.Smooth() { Val = false });

                series.Append(new C.XValues(numberLiteralX));
                series.Append(new C.YValues(numberLiteralY));

                scatterChart.Append(series);

                seriesIndex++;
            }

            addAxisIds(scatterChart, categoryAxisId, valueAxisId);

            plotArea.Append(scatterChart);

            addValueAxis(chartDefinition, chartNode, plotArea, valueAxisId, categoryAxisId, position: AxisPositionValues.Bottom);
            addValueAxis(chartDefinition, chartNode, plotArea, categoryAxisId, valueAxisId, position: AxisPositionValues.Left);

            return GetChartCommon(chartDefinition, chartNode, plotArea);
        }
        private static C.Chart getBubbleChart(ChartDefinition chartDefinition, XElement chartNode)
        {
            uint categoryAxisId = getSafeId();
            uint valueAxisId = getSafeId();

            C.PlotArea plotArea = new();
            plotArea.Append(new Layout());

            BubbleChart bubbleChart = new(
                new VaryColors() { Val = chartDefinition.VaryColors ?? false }
            );

            uint seriesIndex = 0;

            foreach (XElement seriesNode in chartNode.Elements("series"))
            {
                BubbleChartSeries series = new BubbleChartSeries(
                    new C.Index() { Val = seriesIndex },
                    new Order() { Val = seriesIndex },
                    new C.SeriesText(new C.NumericValue(seriesNode.Attribute("name")?.Value ?? ""))
                );

                List<XElement> points = seriesNode.Elements("point").ToList();
                uint pointCount = (uint)points.Count;

                NumberLiteral numberLiteralX = new NumberLiteral();
                numberLiteralX.Append(new C.PointCount() { Val = (uint)pointCount });
                for (int i = 0; i < pointCount; i++)
                {
                    string? x = points[i].Attribute("x")?.Value;

                    numberLiteralX.Append(new C.NumericPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new C.NumericValue(x)
                    });
                }

                NumberLiteral numberLiteralY = new NumberLiteral();
                numberLiteralY.Append(new C.PointCount() { Val = (uint)pointCount });
                for (int i = 0; i < pointCount; i++)
                {
                    string? y = points[i].Attribute("y")?.Value;

                    numberLiteralY.Append(new C.NumericPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new C.NumericValue(y)
                    });
                }

                NumberLiteral numberLiteralSize = new NumberLiteral();
                numberLiteralSize.Append(new C.PointCount() { Val = (uint)pointCount });
                for (int i = 0; i < pointCount; i++)
                {
                    string? size = points[i].Attribute("size")?.Value;

                    numberLiteralSize.Append(new C.NumericPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new C.NumericValue(size)
                    });
                }

                series.Append(new C.XValues(numberLiteralX));
                series.Append(new C.YValues(numberLiteralY));
                series.Append(new C.BubbleSize(numberLiteralSize));

                bubbleChart.Append(series);

                seriesIndex++;
            }

            addAxisIds(bubbleChart, categoryAxisId, valueAxisId);

            bubbleChart.Append(new C.BubbleScale() { Val = (UInt32Value)(uint)clamp(chartDefinition.BubbleScale ?? 100, 0, 300) });

            bubbleChart.Append(new C.ShowNegativeBubbles() { Val = true });

            plotArea.Append(bubbleChart);

            addValueAxis(chartDefinition, chartNode, plotArea, valueAxisId, categoryAxisId, position: AxisPositionValues.Bottom);
            addValueAxis(chartDefinition, chartNode, plotArea, categoryAxisId, valueAxisId, position: AxisPositionValues.Left);

            return GetChartCommon(chartDefinition, chartNode, plotArea);

        }
        private static C.Chart getRadarChart(ChartDefinition chartDefinition, XElement chartNode)
        {
            uint categoryAxisId = getSafeId();
            uint valueAxisId = getSafeId();

            C.PlotArea plotArea = new();
            plotArea.Append(new Layout());

            RadarChart radarChart = new(
                new C.RadarStyle() { Val = mapRadarStyle(chartDefinition.RadarStyle) },
                new C.VaryColors() { Val = chartDefinition.VaryColors ?? false }
            );

            uint seriesIndex = 0;

            foreach (XElement seriesNode in chartNode.Elements("series"))
            {
                RadarChartSeries series = new RadarChartSeries(
                    new C.Index() { Val = seriesIndex },
                    new Order() { Val = seriesIndex },
                    new C.SeriesText(new C.NumericValue(seriesNode.Attribute("name")?.Value ?? ""))
                );

                List<XElement> points = seriesNode.Elements("point").ToList();
                uint pointCount = (uint)points.Count;

                CategoryAxisData catAxisData = new CategoryAxisData();
                C.Values values = new C.Values();

                StringLiteral stringLiteral = new StringLiteral();
                NumberLiteral numberLiteral = new NumberLiteral();

                stringLiteral.Append(new PointCount() { Val = (uint)pointCount });

                for (int i = 0; i < pointCount; i++)
                {
                    string categoryName = points[i].Attribute("category")?.Value ?? $"Kategori {i + 1}";
                    string valueStr = points[i].Attribute("value")?.Value ?? "0";

                    stringLiteral.Append(new StringPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new C.NumericValue(categoryName)
                    });
                    numberLiteral.Append(new NumericPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new C.NumericValue(valueStr)
                    });
                }

                catAxisData.Append(stringLiteral);
                values.Append(numberLiteral);

                series.Append(catAxisData);
                series.Append(values);

                radarChart.Append(series);

                seriesIndex++;
            }

            addAxisIds(radarChart, categoryAxisId, valueAxisId);

            plotArea.Append(radarChart);

            addCategoryAxis(chartDefinition, chartNode, plotArea, categoryAxisId, valueAxisId, position: AxisPositionValues.Bottom);
            addValueAxis(chartDefinition, chartNode, plotArea, categoryAxisId, valueAxisId, position: AxisPositionValues.Left);

            C.Chart chart = new C.Chart();

            string chartTitle = chartNode.Attribute("title")?.Value ?? "";

            A.Text text = new A.Text(chartTitle);
            A.Run run = new A.Run(text);
            A.Paragraph paragraph = new A.Paragraph(run);
            C.RichText richText = new C.RichText(new A.BodyProperties(), new A.ListStyle(), paragraph);
            C.ChartText chartText = new C.ChartText(richText);
            C.Title title = new C.Title(chartText, new Overlay() { Val = false });
            Overlay overlay = new Overlay() { Val = false };

            chart.Append(title);

            chart.Append(plotArea);

            chart.Append(new DisplayBlanksAs() { Val = mapBlanksDisplayedAs(chartDefinition.BlanksDisplayedAs ?? BlanksDisplayedAs.Gap) });

            chart.Append(new AutoTitleDeleted() { Val = !chartDefinition.AutoTitle ?? false });

            chart.Append(new PlotVisibleOnly() { Val = chartDefinition.PlotVisibleOnly ?? false });

            chart.Append(new ShowDataLabelsOverMaximum() { Val = chartDefinition.ShowDataLabelsOverMaximum ?? true });


            return chart;

        }
        private static int clamp(int v, int min, int max)
        {
            if (v < min) return min;
            if (v > max) return max;

            return v;
        }
        private static uint getSafeId()
        {
            byte[] guidBytes = Guid.NewGuid().ToByteArray();
            return BitConverter.ToUInt32(guidBytes, 0) & 0x7FFFFFFF;
        }
        private static void addAxisIds(OpenXmlCompositeElement owner, params uint[] ids)
        {
            if (owner is null || ids is null)
                return;

            foreach (uint id in ids)
                owner.Append(new C.AxisId { Val = id });
        }
        private static void addValueAxis(ChartDefinition chartDefinition, XElement chartNode, C.PlotArea plotArea, uint categoryAxisId, uint valueAxisId, AxisPositionValues? position = null)
        {
            AxisPositionValues axisPos = position ?? AxisPositionValues.Left;

            ValueAxis valAx = new ValueAxis(
                new C.AxisId() { Val = valueAxisId },
                new Scaling(new Orientation() { Val = C.OrientationValues.MinMax }),
                new Delete() { Val = !chartDefinition.ShowValueAxis },
                new AxisPosition() { Val = axisPos },
                new C.MajorTickMark() { Val = chartDefinition.Type == ChartType.Radar ? C.TickMarkValues.Cross : C.TickMarkValues.Outside },
                new C.MinorTickMark() { Val = C.TickMarkValues.None },
                new C.NumberingFormat() { FormatCode = "General", SourceLinked = true },
                new TickLabelPosition() { Val = TickLabelPositionValues.NextTo },
                new CrossingAxis() { Val = categoryAxisId },
                new Crosses() { Val = CrossesValues.AutoZero },
                new CrossBetween() { Val = CrossBetweenValues.Between }
            );

            plotArea.Append(valAx);
        }
        private static void addCategoryAxis(ChartDefinition chartDefinition, XElement chartNode, C.PlotArea plotArea, uint categoryAxisId, uint valueAxisId, AxisPositionValues? position = null)
        {
            AxisPositionValues axisPos = position ?? AxisPositionValues.Bottom;

            CategoryAxis catAxis = new CategoryAxis(
                new C.AxisId() { Val = categoryAxisId },
                new Delete() { Val = !chartDefinition.ShowCategoryAxis },
                new Scaling(new Orientation() { Val = C.OrientationValues.MinMax }),
                new AxisPosition() { Val = axisPos },
                new TickLabelPosition() { Val = TickLabelPositionValues.NextTo },
                new CrossingAxis() { Val = valueAxisId },
                new Crosses() { Val = CrossesValues.AutoZero },
                new AutoLabeled() { Val = true },
                new LabelAlignment() { Val = LabelAlignmentValues.Center },
                new LabelOffset() { Val = 100 },
                new C.TickLabelSkip() { Val = 1 },
                new C.TickMarkSkip() { Val = 1 },
                new C.NoMultiLevelLabels() { Val = true }
            );

            switch (chartDefinition.Type)
            {
                case ChartType.Bar:
                case ChartType.Column:
                case ChartType.Line:
                case ChartType.Area:
                case ChartType.Radar:
                    catAxis.InsertAt(new C.MajorTickMark() { Val = C.TickMarkValues.None }, 3);
                    catAxis.InsertAt(new C.MinorTickMark() { Val = C.TickMarkValues.Outside }, 4);
                    break;
                case ChartType.Scatter:
                case ChartType.Bubble:
                    catAxis.InsertAt(new C.MajorTickMark() { Val = C.TickMarkValues.Outside }, 3);
                    catAxis.InsertAt(new C.MinorTickMark() { Val = C.TickMarkValues.None }, 4);
                    break;
                case ChartType.Pie:
                case ChartType.Doughnut:
                default:
                    catAxis.InsertAt(new C.MajorTickMark() { Val = C.TickMarkValues.None }, 3);
                    catAxis.InsertAt(new C.MinorTickMark() { Val = C.TickMarkValues.None }, 4);
                    break;
            }

            applyAxisDefinition(catAxis, chartDefinition.CategoryAxis);

            plotArea.Append(catAxis);
        }
        private static void applyAxisDefinition(OpenXmlCompositeElement catAxis, AxisDefinition axisDefinition)
        {
            if (axisDefinition is null || catAxis is null)
                return;

            Scaling scaling = catAxis.Elements<Scaling>().FirstOrDefault();
            if (scaling is null)
            {
                scaling = new Scaling();
                catAxis.PrependChild(scaling);
            }

            if (axisDefinition.Min.HasValue)
                scaling.Append(new C.MinAxisValue() { Val = axisDefinition.Min.Value });

            if (axisDefinition.Max.HasValue && (!axisDefinition.Min.HasValue || axisDefinition.Min.Value <= axisDefinition.Max.Value))
                scaling.Append(new C.MaxAxisValue() { Val = axisDefinition.Max.Value });

            if (axisDefinition.MajorUnit.HasValue && catAxis is ValueAxis)
                catAxis.Append(new MajorUnit() { Val = axisDefinition.MajorUnit.Value });

            C.ChartShapeProperties chartShapeProperties = new C.ChartShapeProperties();

            if (axisDefinition.AxisLineFormat?.FillDefinition?.SolidFillDefinition?.Color is not null)
            {
                string colorValue = axisDefinition.AxisLineFormat.FillDefinition.SolidFillDefinition.Color.Color;
                if (colorValue.StartsWith("#"))
                    colorValue = colorValue.Substring(1);

                A.RgbColorModelHex colorHex = new A.RgbColorModelHex() { Val = colorValue };

                if (axisDefinition.AxisLineFormat.FillDefinition.SolidFillDefinition.Color.Transparency.HasValue)
                {
                    int transparency = axisDefinition.AxisLineFormat.FillDefinition.SolidFillDefinition.Color.Transparency.Value;
                    int alphaValue = (int)Math.Round((100 - transparency) * 1000.0);
                    int finalAlpha = Math.Max(0, Math.Min(100000, alphaValue));

                    Console.WriteLine($"Debug Axis: transparency={transparency}, alphaValue={alphaValue}, finalAlpha={finalAlpha}");

                    colorHex.Append(new A.Alpha() { Val = finalAlpha });
                }

                A.SolidFill solidFill = new A.SolidFill(colorHex);
                chartShapeProperties.Append(solidFill);
                catAxis.Append(chartShapeProperties);
            }

            if (axisDefinition.AxisLineFormat?.FillDefinition?.PatternFillDefinition is not null)
            {
                PatternDefinition patternDef = axisDefinition.AxisLineFormat.FillDefinition.PatternFillDefinition;

                A.PatternFill patternFill = new A.PatternFill();

                if (patternDef.Preset != null)
                {
                    A.PresetPatternValues presetValue = mapPatternPreset(patternDef.Preset);
                    patternFill.Preset = presetValue;
                }

                if (patternDef.ForegroundColor != null)
                {
                    A.ForegroundColor fgColor = new A.ForegroundColor();
                    string fgColorValue = patternDef.ForegroundColor.Color;
                    if (fgColorValue.StartsWith("#"))
                        fgColorValue = fgColorValue.Substring(1);

                    A.RgbColorModelHex fgColorHex = new A.RgbColorModelHex() { Val = fgColorValue };

                    if (patternDef.ForegroundColor.Transparency.HasValue)
                    {
                        int transparency = patternDef.ForegroundColor.Transparency.Value;
                        int alphaValue = (int)Math.Round((100 - transparency) * 1000.0);
                        int finalAlpha = Math.Max(0, Math.Min(100000, alphaValue));
                        fgColorHex.Append(new A.Alpha() { Val = finalAlpha });
                    }

                    fgColor.Append(fgColorHex);
                    patternFill.Append(fgColor);
                }

                if (patternDef.BackgroundColor != null)
                {
                    A.BackgroundColor bgColor = new A.BackgroundColor();
                    string bgColorValue = patternDef.BackgroundColor.Color;
                    if (bgColorValue.StartsWith("#"))
                        bgColorValue = bgColorValue.Substring(1);

                    A.RgbColorModelHex bgColorHex = new A.RgbColorModelHex() { Val = bgColorValue };

                    if (patternDef.BackgroundColor.Transparency.HasValue)
                    {
                        int transparency = patternDef.BackgroundColor.Transparency.Value;
                        int alphaValue = (int)Math.Round((100 - transparency) * 1000.0);
                        int finalAlpha = Math.Max(0, Math.Min(100000, alphaValue));
                        bgColorHex.Append(new A.Alpha() { Val = finalAlpha });
                    }

                    bgColor.Append(bgColorHex);
                    patternFill.Append(bgColor);
                }

                chartShapeProperties.Append(patternFill);
                catAxis.Append(chartShapeProperties);
            }

            if (axisDefinition.AxisLineFormat?.FillDefinition?.GradientFillDefinition is not null)
            {
                GradientDefinition gradientDefinition = axisDefinition.AxisLineFormat.FillDefinition.GradientFillDefinition;
                int angleDegrees = clamp(gradientDefinition.Angle ?? 45, 0, 360);

                A.LinearGradientFill linearGradient = new A.LinearGradientFill()
                {
                    Angle = new Int32Value(angleDegrees * 60000),
                    Scaled = gradientDefinition.Scaled ?? true
                };

                //A.GradientStopList gradientStopList = new A.GradientStopList();

                //foreach (GradientStopDefinition stop in gradientDefinition.Stops)
                //{
                //    if (stop.Color != null)
                //    {
                //        A.GradientStop gradientStop = new A.GradientStop()
                //        {
                //            Position = new Int32Value((int)Math.Round(stop.Position * 1000.0))
                //        };

                //        string colorValue = stop.Color.Color;
                //        if (colorValue.StartsWith("#"))
                //            colorValue = colorValue.Substring(1);

                //        A.RgbColorModelHex stopColorHex = new A.RgbColorModelHex() { Val = colorValue };

                //        if (stop.Color.Transparency.HasValue)
                //        {
                //            int transparency = stop.Color.Transparency.Value;
                //            int alphaValue = (int)Math.Round((100 - transparency) * 1000.0);
                //            int finalAlpha = Math.Max(0, Math.Min(100000, alphaValue));
                //            stopColorHex.Append(new A.Alpha() { Val = finalAlpha });
                //        }

                //        gradientStop.Append(stopColorHex);
                //        gradientStopList.Append(gradientStop);
                //    }
                //}

                //linearGradient.Append(gradientStopList);
                chartShapeProperties.Append(linearGradient);
                catAxis.Append(chartShapeProperties);
            }

            if (axisDefinition.AxisLineFormat?.LineDefinition is not null)
            {
                LineDefinition lineDefinition = axisDefinition.AxisLineFormat.LineDefinition;

                A.Outline outline = new A.Outline();


                if (!lineDefinition.Visible)
                {
                    outline.RemoveAllChildren();
                    outline.Append(new A.NoFill());
                }
                else
                {
                    if (lineDefinition.Width is not null)
                        outline.Width = lineDefinition.Width.ToEmu();

                    if (lineDefinition.DashPreset is not null)
                        outline.Append(new A.PresetDash() { Val = mapDash(lineDefinition.DashPreset.Value) });

                    if (lineDefinition.CompoundPreset is not null)
                        outline.CompoundLineType = mapCompound(lineDefinition.CompoundPreset.Value);

                    if (lineDefinition.CapPreset is not null)
                        outline.CapType = mapLineCap(lineDefinition.CapPreset.Value);

                    if (lineDefinition?.JoinPreset is null)
                        return;

                    switch (lineDefinition.JoinPreset.Value)
                    {
                        case LineJoinPreset.Round:
                            outline.Append(new A.Round());
                            break;
                        case LineJoinPreset.Bevel:
                            outline.Append(new A.Bevel());
                            break;
                        case LineJoinPreset.Miter:
                            outline.Append(new A.Miter() { Limit = (lineDefinition.JoinMiterLimit ?? 800000) });
                            break;
                        default:
                            break;
                    }

                    if (lineDefinition.Begin_Arrow_Type is not null || lineDefinition.Begin_Arrow_Width is not null || lineDefinition.Begin_Arrow_Length is not null)
                    {
                        outline.RemoveAllChildren<A.HeadEnd>();
                        outline.Append(new A.HeadEnd
                        {
                            Type = lineDefinition.Begin_Arrow_Type is null ? null : mapLineEndType(lineDefinition.Begin_Arrow_Type.Value),
                            Width = lineDefinition.Begin_Arrow_Width is null ? null : mapLineEndWidth(lineDefinition.Begin_Arrow_Width.Value),
                            Length = lineDefinition.Begin_Arrow_Length is null ? null : mapLineEndLength(lineDefinition.Begin_Arrow_Length.Value),
                        });
                    }

                    if (lineDefinition.End_Arrow_Type is not null || lineDefinition.End_Arrow_Width is not null || lineDefinition.End_Arrow_Length is not null)
                    {
                        outline.RemoveAllChildren<A.TailEnd>();
                        outline.Append(new A.TailEnd
                        {
                            Type = lineDefinition.End_Arrow_Type is null ? null : mapLineEndType(lineDefinition.End_Arrow_Type.Value),
                            Width = lineDefinition.End_Arrow_Width is null ? null : mapLineEndWidth(lineDefinition.End_Arrow_Width.Value),
                            Length = lineDefinition.End_Arrow_Length is null ? null : mapLineEndLength(lineDefinition.End_Arrow_Length.Value),
                        });
                    }

                }

                chartShapeProperties.Append(outline);
                catAxis.Append(chartShapeProperties);
            }


        }




        private static BarGroupingValues mapBarGrouping(GroupingType value)
        {
            return value switch
            {
                GroupingType.Clustered => BarGroupingValues.Clustered,
                GroupingType.Stacked => BarGroupingValues.Stacked,
                GroupingType.PercentStacked => BarGroupingValues.PercentStacked,
                GroupingType.Standard => BarGroupingValues.Standard,
                _ => BarGroupingValues.Clustered
            };
        }
        private static ScatterStyleValues mapScatterStyle(ScatterStyle value)
        {
            return value switch
            {
                ScatterStyle.Line => ScatterStyleValues.Line,
                ScatterStyle.LineMarker => ScatterStyleValues.LineMarker,
                ScatterStyle.Marker => ScatterStyleValues.Marker,
                ScatterStyle.Smooth => ScatterStyleValues.Smooth,
                ScatterStyle.SmoothMarker => ScatterStyleValues.SmoothMarker,
                _ => ScatterStyleValues.Marker
            };
        }
        private static C.RadarStyleValues mapRadarStyle(RadarStyle value)
        {
            return value switch
            {
                RadarStyle.Standard => C.RadarStyleValues.Standard,
                RadarStyle.Marker => C.RadarStyleValues.Marker,
                RadarStyle.Filled => C.RadarStyleValues.Filled,
                _ => C.RadarStyleValues.Standard,
            };
        }
        private static DisplayBlanksAsValues mapBlanksDisplayedAs(BlanksDisplayedAs value)
        {
            return value switch
            {
                BlanksDisplayedAs.Gap => DisplayBlanksAsValues.Gap,
                BlanksDisplayedAs.Zero => DisplayBlanksAsValues.Zero,
                BlanksDisplayedAs.Span => DisplayBlanksAsValues.Span,
                _ => DisplayBlanksAsValues.Gap
            };
        }
        private static A.PresetPatternValues mapPatternPreset(PatternFillPreset value)
        {
            return value switch
            {
                PatternFillPreset.Percent5 => A.PresetPatternValues.Percent5,
                PatternFillPreset.Percent10 => A.PresetPatternValues.Percent10,
                PatternFillPreset.Percent20 => A.PresetPatternValues.Percent20,
                PatternFillPreset.Percent25 => A.PresetPatternValues.Percent25,
                PatternFillPreset.Percent30 => A.PresetPatternValues.Percent30,
                PatternFillPreset.Percent40 => A.PresetPatternValues.Percent40,
                PatternFillPreset.Percent50 => A.PresetPatternValues.Percent50,
                PatternFillPreset.Percent60 => A.PresetPatternValues.Percent60,
                PatternFillPreset.Percent70 => A.PresetPatternValues.Percent70,
                PatternFillPreset.Percent75 => A.PresetPatternValues.Percent75,
                PatternFillPreset.Percent80 => A.PresetPatternValues.Percent80,
                PatternFillPreset.Percent90 => A.PresetPatternValues.Percent90,
                PatternFillPreset.Horizontal => A.PresetPatternValues.Horizontal,
                PatternFillPreset.Vertical => A.PresetPatternValues.Vertical,
                PatternFillPreset.LightHorizontal => A.PresetPatternValues.LightHorizontal,
                PatternFillPreset.LightVertical => A.PresetPatternValues.LightVertical,
                PatternFillPreset.DarkHorizontal => A.PresetPatternValues.DarkHorizontal,
                PatternFillPreset.DarkVertical => A.PresetPatternValues.DarkVertical,
                PatternFillPreset.NarrowHorizontal => A.PresetPatternValues.NarrowHorizontal,
                PatternFillPreset.NarrowVertical => A.PresetPatternValues.NarrowVertical,
                PatternFillPreset.DashedHorizontal => A.PresetPatternValues.DashedHorizontal,
                PatternFillPreset.DashedVertical => A.PresetPatternValues.DashedVertical,
                PatternFillPreset.Cross => A.PresetPatternValues.Cross,
                PatternFillPreset.DownwardDiagonal => A.PresetPatternValues.DownwardDiagonal,
                PatternFillPreset.UpwardDiagonal => A.PresetPatternValues.UpwardDiagonal,
                PatternFillPreset.LightDownwardDiagonal => A.PresetPatternValues.LightDownwardDiagonal,
                PatternFillPreset.LightUpwardDiagonal => A.PresetPatternValues.LightUpwardDiagonal,
                PatternFillPreset.DarkDownwardDiagonal => A.PresetPatternValues.DarkDownwardDiagonal,
                PatternFillPreset.DarkUpwardDiagonal => A.PresetPatternValues.DarkUpwardDiagonal,
                PatternFillPreset.WideDownwardDiagonal => A.PresetPatternValues.WideDownwardDiagonal,
                PatternFillPreset.WideUpwardDiagonal => A.PresetPatternValues.WideUpwardDiagonal,
                PatternFillPreset.DashedDownwardDiagonal => A.PresetPatternValues.DashedDownwardDiagonal,
                PatternFillPreset.DashedUpwardDiagonal => A.PresetPatternValues.DashedUpwardDiagonal,
                PatternFillPreset.DiagonalCross => A.PresetPatternValues.DiagonalCross,
                PatternFillPreset.SmallCheck => A.PresetPatternValues.SmallCheck,
                PatternFillPreset.LargeCheck => A.PresetPatternValues.LargeCheck,
                PatternFillPreset.SmallGrid => A.PresetPatternValues.SmallGrid,
                PatternFillPreset.LargeGrid => A.PresetPatternValues.LargeGrid,
                PatternFillPreset.DotGrid => A.PresetPatternValues.DotGrid,
                PatternFillPreset.SmallConfetti => A.PresetPatternValues.SmallConfetti,
                PatternFillPreset.LargeConfetti => A.PresetPatternValues.LargeConfetti,
                PatternFillPreset.HorizontalBrick => A.PresetPatternValues.HorizontalBrick,
                PatternFillPreset.DiagonalBrick => A.PresetPatternValues.DiagonalBrick,
                PatternFillPreset.SolidDiamond => A.PresetPatternValues.SolidDiamond,
                PatternFillPreset.OpenDiamond => A.PresetPatternValues.OpenDiamond,
                PatternFillPreset.DottedDiamond => A.PresetPatternValues.DottedDiamond,
                PatternFillPreset.Plaid => A.PresetPatternValues.Plaid,
                PatternFillPreset.Sphere => A.PresetPatternValues.Sphere,
                PatternFillPreset.Weave => A.PresetPatternValues.Weave,
                PatternFillPreset.Divot => A.PresetPatternValues.Divot,
                PatternFillPreset.Shingle => A.PresetPatternValues.Shingle,
                PatternFillPreset.Wave => A.PresetPatternValues.Wave,
                PatternFillPreset.Trellis => A.PresetPatternValues.Trellis,
                PatternFillPreset.ZigZag => A.PresetPatternValues.ZigZag,
                _ => A.PresetPatternValues.DotGrid
            };
        }
        private static A.PresetLineDashValues mapDash(DashPreset value)
        {
            return value switch
            {
                DashPreset.Solid => A.PresetLineDashValues.Solid,
                DashPreset.Dot => A.PresetLineDashValues.Dot,
                DashPreset.Dash => A.PresetLineDashValues.Dash,
                DashPreset.LargeDash => A.PresetLineDashValues.LargeDash,
                DashPreset.DashDot => A.PresetLineDashValues.DashDot,
                DashPreset.LargeDashDot => A.PresetLineDashValues.LargeDashDot,
                DashPreset.LargeDashDotDot => A.PresetLineDashValues.LargeDashDotDot,
                DashPreset.SystemDash => A.PresetLineDashValues.SystemDash,
                DashPreset.SystemDot => A.PresetLineDashValues.SystemDot,
                DashPreset.SystemDashDot => A.PresetLineDashValues.SystemDashDot,
                DashPreset.SystemDashDotDot => A.PresetLineDashValues.SystemDashDotDot,
                _ => A.PresetLineDashValues.Solid
            };
        }
        private static A.CompoundLineValues mapCompound(CompoundLinePreset value)
        {
            return value switch
            {
                CompoundLinePreset.Single => A.CompoundLineValues.Single,
                CompoundLinePreset.Double => A.CompoundLineValues.Double,
                CompoundLinePreset.ThickThin => A.CompoundLineValues.ThickThin,
                CompoundLinePreset.ThinThick => A.CompoundLineValues.ThinThick,
                CompoundLinePreset.Triple => A.CompoundLineValues.Triple,
                _ => A.CompoundLineValues.Single
            };
        }
        private static A.LineCapValues mapLineCap(LineCapPreset value)
        {
            return value switch
            {
                LineCapPreset.Flat => A.LineCapValues.Flat,
                LineCapPreset.Round => A.LineCapValues.Round,
                LineCapPreset.Square => A.LineCapValues.Square,
                _ => A.LineCapValues.Flat
            };
        }
        private static A.LineEndValues mapLineEndType(LineEndPreset value)
        {
            return value switch
            {
                LineEndPreset.None => A.LineEndValues.None,
                LineEndPreset.Triangle => A.LineEndValues.Triangle,
                LineEndPreset.Stealth => A.LineEndValues.Stealth,
                LineEndPreset.Diamond => A.LineEndValues.Diamond,
                LineEndPreset.Oval => A.LineEndValues.Oval,
                LineEndPreset.Arrow => A.LineEndValues.Arrow,
                _ => A.LineEndValues.None
            };
        }
        private static A.LineEndWidthValues mapLineEndWidth(LineEndWidthPreset value)
        {
            return value switch
            {
                LineEndWidthPreset.Small => A.LineEndWidthValues.Small,
                LineEndWidthPreset.Medium => A.LineEndWidthValues.Medium,
                LineEndWidthPreset.Large => A.LineEndWidthValues.Large,
                _ => A.LineEndWidthValues.Medium
            };
        }
        private static A.LineEndLengthValues mapLineEndLength(LineEndLengthPreset value)
        {
            return value switch
            {
                LineEndLengthPreset.Small => A.LineEndLengthValues.Small,
                LineEndLengthPreset.Medium => A.LineEndLengthValues.Medium,
                LineEndLengthPreset.Large => A.LineEndLengthValues.Large,
                _ => A.LineEndLengthValues.Medium
            };
        }
    }

    class SeriesDefinition
    {
        public TitleDefinition Title { get; set; }
        public ColorA ColorA { get { return Format?.FillDefinition?.SolidFillDefinition?.Color; } }
        public string Color { get { return ColorA?.Color; } }
        public int? ColorTransparency { get { return ColorA?.Transparency; } }
        public List<PointDefinition> Points { get; set; } = [];
        public FormatDefinition Format { get; set; }
        public MarkerDefinition Marker { get; set; }
        public ValueLabelDefinition ValueLabels { get; set; } //<value-labels>
    }

    class PointDefinition
    {
        // Categorical charts
        public string Category { get; set; }
        public double? Value { get; set; }

        // Scatter / Bubble
        public double? X { get; set; }
        public double? Y { get; set; }

        // Bubble only
        public double? Size { get; set; }

        //UI
        public bool HasStyle { get { return PointFormat is not null || MarkerType is not null || MarkerSize is not null; } }
        public MarkerType? MarkerType { get; set; } // circle/square/…
        public int? MarkerSize { get; set; } // 2..72 (pt karşılığı byte range: 2..72) 
        public FormatDefinition PointFormat { get; set; }
        public ValueLabelDefinition ValueLabels { get; set; }
        public SeriesDefinition SeriesDefinition { get; set; }

    }
    class AxisDefinition
    {
        public TitleDefinition Title { get; set; } // <title>
        public FormatDefinition AxisLineFormat { get; set; } // <format>
        public TextFormatDefinition TickLabelTextFormat { get; set; } // <text-format>

        // Only valid for numeric axes
        public double? Min { get; set; }
        public double? Max { get; set; }
        public double? MajorUnit { get; set; }
    }
    class LegendDefinition
    {
        public bool Show { get; set; }
        public LegendPosition Position { get; set; }
        public FormatDefinition BoxFormat { get; set; }
        public TextFormatDefinition LegendEntryTextFormat { get; set; }
    }
    class GridDefinition
    {
        public bool ShowHorizontal { get; set; } = true;
        public bool ShowVertical { get; set; } = true;
        public FormatDefinition HorizontalFormat { get; set; }
        public FormatDefinition VerticalFormat { get; set; }
    }
    class ValueLabelDefinition
    {
        public DataLabelPosition? Position { get; set; }
        public bool Show { get; set; }
        public bool ShowLegendKey { get; set; }
        public bool ShowCategoryName { get; set; }
        public bool ShowSeriesName { get; set; }
        public bool ShowPercent { get; set; }
        public bool ShowBubbleSize { get; set; }
        public TextFormatDefinition TextFormat { get; set; }
    }
    class ChartLayoutDefinition
    {
        public bool HasManualLayout => X.HasValue || Y.HasValue || Width.HasValue || Height.HasValue;
        public int? X { get; set; }
        public int? Y { get; set; }
        public int? Width { get; set; }
        public int? Height { get; set; }
        public LayoutMode? XMode { get; set; }
        public LayoutMode? YMode { get; set; }
        public LayoutTarget? Target { get; set; }
    }
    class ThreeDViewDefinition
    {
        //  <view-3d > <floor-format /> <back-wall-format /> <side-wall-format /> </view-3d>
        // View3D ortak
        public int? RotationX { get; set; } // -90..90
        public int? RotationY { get; set; } // 0..360
        public bool? RightAngleAxes { get; set; } // true -> Perspective yok sayılır
        public int? Perspective { get; set; } // 0..240 (RightAngleAxes=false iken)
        public int? DepthPercent { get; set; } // 20..2000
        public int? HeightPercent { get; set; } // 5..500

        // Bar/Column 3D özel
        public int? GapWidth { get; set; } // 0..500
        public int? GapDepth { get; set; } // 0..500
        public ThreeDShape? Shape { get; set; } // box, cylinder, cone, cone-to-max, pyramid, pyramid-to-max

        // Duvar/zemin (opsiyonel)
        public bool? ShowFloor { get; set; }
        public bool? ShowBackWall { get; set; }
        public bool? ShowSideWall { get; set; }

        //Fill & Border
        public FormatDefinition DefaultFormat { get; set; }
        public FormatDefinition FloorFormat { get; set; }
        public FormatDefinition BackWallFormat { get; set; }
        public FormatDefinition SideWallFormat { get; set; }
    }
    class MarkerDefinition
    {
        public MarkerType? Type { get; set; } // circle/square/…
        public int? Size { get; set; } // 2..72 (pt karşılığı byte range: 2..72)
        public FormatDefinition Format { get; set; } // marker/c:spPr (fill/line/effects)
    }
    class TitleDefinition
    {
        public string Text { get; set; }
        public TextFormatDefinition TextFormat { get; set; }

        public static implicit operator string(TitleDefinition d) => d?.Text;
    }
    class DataTableDefinition
    {
        public bool Show { get; set; }
        public bool ShowHorizontalBorder { get; set; }
        public bool ShowVerticalBorder { get; set; }
        public bool ShowOutlineBorder { get; set; }
        public bool ShowLegendKey { get; set; } // legend keys ilk kolonda
        public FormatDefinition BoxFormat { get; set; } // c:spPr
        public TextFormatDefinition TextFormat { get; set; } // c:txPr
    }

    class FormatDefinition
    {
        public FillDefinition FillDefinition { get; set; }
        public LineDefinition LineDefinition { get; set; }
        public EffectsDefinition EffectsDefinition { get; set; }
    }
    class TextFormatDefinition
    {
        public HorizontalAlign? AlignHorizontal { get; set; }
        public VerticalAlign? AlignVertical { get; set; }
        public TextWrapPreset? Wrap { get; set; }
        public Dimension MarginLeft { get; set; }
        public Dimension MarginRight { get; set; }
        public Dimension MarginTop { get; set; }
        public Dimension MarginBottom { get; set; }
        public int? Rotate { get; set; }
        public FillDefinition TextFillDefinition { get; set; }
        public LineDefinition TextOutlineDefinition { get; set; }
        public EffectsDefinition EffectsDefinition { get; set; }
        public FontDefinition FontDefinition { get; set; }
    }

    class FillDefinition
    {
        public SolidDefinition SolidFillDefinition { get; set; }
        public PatternDefinition PatternFillDefinition { get; set; }
        public GradientDefinition GradientFillDefinition { get; set; }
    }
    class SolidDefinition
    {
        public ColorA Color { get; set; }
    }
    class PatternDefinition
    {
        public PatternFillPreset Preset { get; set; }
        public ColorA ForegroundColor { get; set; }
        public ColorA BackgroundColor { get; set; }
    }
    class GradientDefinition
    {
        public int? Angle { get; set; }
        public bool? Scaled { get; set; }
        public List<GradientStopDefinition> Stops { get; set; } = [];
    }
    class GradientStopDefinition
    {
        public int Position { get; set; }
        public ColorA Color { get; set; }
    }
    class LineDefinition
    {
        public bool Visible { get; set; }
        public Dimension Width { get; set; }
        public SolidDefinition SolidLineDefinition { get; set; }
        public GradientDefinition GradientLineDefinition { get; set; }
        public DashPreset? DashPreset { get; set; }
        public CompoundLinePreset? CompoundPreset { get; set; }
        public LineCapPreset? CapPreset { get; set; }
        public LineJoinPreset? JoinPreset { get; set; }
        public int? JoinMiterLimit { get; set; }
        public LineEndPreset? Begin_Arrow_Type { get; set; }
        public LineEndWidthPreset? Begin_Arrow_Width { get; set; }
        public LineEndLengthPreset? Begin_Arrow_Length { get; set; }
        public LineEndPreset? End_Arrow_Type { get; set; }
        public LineEndWidthPreset? End_Arrow_Width { get; set; }
        public LineEndLengthPreset? End_Arrow_Length { get; set; }
    }

    class EffectsDefinition
    {
        //   <effects> ...</effects>      
        public ShadowDefinition ShadowDefinition { get; set; }
        public GlowDefinition GlowDefinition { get; set; }
        public SoftEdgesDefinition SoftEdgesDefinition { get; set; }
        public Format3DDefinition Format3dDefinition { get; set; }
        public ReflectionDefinition ReflectionDefinition { get; set; }
    }
    class ShadowDefinition
    {
        //   <shadow type="inner/outer/perspective/preset" preset="A.PresetShadowValues.*" color="#dedede" transparency="80" blur-radius="5pt" angle="45" distance="4pt" />

        public ShadowType? Type { get; set; }
        public ShadowPreset? Preset { get; set; }
        public ColorA Color { get; set; }
        public Dimension BlurRadius { get; set; }
        public Dimension Distance { get; set; }
        public int? Angle { get; set; }
    }
    class GlowDefinition
    {
        //   <glow color="#RRGGBB" transparency="0..100" size="pt"/>
        public ColorA Color { get; set; }
        public Dimension Size { get; set; }
    }
    class SoftEdgesDefinition
    {
        //   <soft-edges size="pt"/>

        public Dimension Size { get; set; }
    }
    class ReflectionDefinition
    {
        //    <reflection blur-radius="pt" distance="pt" start-transparency="0..100" end-transparency="0..100" start-position="0..100" end-position="0..100"/>
        public bool HasData
        {
            get
            {
                return Blur is not null
                       || Distance is not null
                       || StartTransparency is int
                       || EndTransparency is int
                       || StartPosition is int
                       || EndPosition is int;
            }
        }
        public Dimension Blur { get; set; }
        public Dimension Distance { get; set; }
        public int? StartTransparency { get; set; }
        public int? EndTransparency { get; set; }
        public int? StartPosition { get; set; }
        public int? EndPosition { get; set; }
    }
    class Format3DDefinition
    {
        //   <format-3d material="plastic|metal|..."> ...</format3d>         

        public MaterialPreset? Material { get; set; }
        public BevelDefinition BevelDefinition { get; set; }
        public LightingDefinition LightingDefinition { get; set; }
    }
    class BevelDefinition
    {
        //   <bevel top-width="pt" top-height="pt" bottom-width="pt" bottom-height="pt" top-preset="relaxedInset|circle|…" bottom-preset="relaxedInset|circle|…"/>
        public bool HasData { get { return TopPreset is not null || BottomPreset is not null; } }
        public Dimension TopWidth { get; set; }
        public Dimension TopHeight { get; set; }
        public Dimension BottomWidth { get; set; }
        public Dimension BottomHeight { get; set; }
        public BevelPreset? TopPreset { get; set; }
        public BevelPreset? BottomPreset { get; set; }
    }
    class LightingDefinition
    {
        // <lighting preset="threePt|soft|harsh|…"  angle="deg"  direction="top|left|..."/>
        public LightingPreset? Preset { get; set; }
        public LightingDirection? Direction { get; set; }
        public int? Angle { get; set; }
    }

    class FontDefinition
    {
        //   <font family="Segoe UI" size="pt" bold="true|false" italic="true|false" underline="none|single|double" strike="none|single|double" caps="none|small|all" kerning="pt" spacing="pt"/>
        public string Family { get; set; }
        public Dimension Size { get; set; }
        public bool? Bold { get; set; }
        public bool? Italic { get; set; }
        public TextUnderlineStyle? Underline { get; set; }
        public TextStrikeStyle? Strike { get; set; }
        public TextCapsStyle? Caps { get; set; }
        public Dimension Kerning { get; set; }
        public Dimension Spacing { get; set; }
    }

    class Dimension
    {
        public double Value { get; set; }
        public UnitType Unit { get; set; }
        public              /*Ctor*/    Dimension(double value, UnitType unit)
        {
            Value = value;
            Unit = unit;
        }
        public int ToEmu()
        {
            return Unit switch
            {
                UnitType.Cm => (int)(Value * 360000),
                UnitType.Pt => (int)(Value * 12700),
                UnitType.In => (int)(Value * 914400),
                UnitType.Px => (int)(Value * 9525),
                _ => throw new InvalidOperationException("Unknown unit")
            };
        }
        public override string ToString() => $"{Value}{Unit.ToString().ToLower()}";
        public static Dimension Parse(string input, bool allowNegative = false)
        {
            if (string.IsNullOrWhiteSpace(input))
                return null;

            Match match = Regex.Match(input.Trim().ToLowerInvariant(), @"^(\d+(?:\.\d+)?)(px|pt|cm|in)?$", RegexOptions.IgnoreCase);

            if (!match.Success)
                return null;

            double value = double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
            string unitStr = match.Groups[2].Success ? match.Groups[2].Value.ToLower() : "px";

            if (value < 0 && !allowNegative)
                return null;

            UnitType unit = unitStr switch
            {
                "cm" => UnitType.Cm,
                "pt" => UnitType.Pt,
                "in" => UnitType.In,
                "px" => UnitType.Px,
                _ => throw new FormatException($"Unknown unit: '{unitStr}'")
            };

            return new Dimension(value, unit);
        }
    }
    class ColorA
    {
        public string Color { get; set; }
        public int? Transparency { get; set; }

        public  /*Ctor*/    ColorA(string color, int? transparency)
        {
            Color = color;
            Transparency = transparency;
        }

        public static ColorA Parse(XElement element, string colorAttr, string transparencyAttr)
        {
            if (element is null)
                return null;

            string? color = element.Attribute(colorAttr)?.Value;

            if (string.IsNullOrWhiteSpace(color))
                return null;

            int? transparency = null;
            string? transparencyStr = element?.Attribute(transparencyAttr)?.Value;

            // Debug: transparency değerini kontrol et
            Console.WriteLine($"Debug: colorAttr='{colorAttr}', transparencyAttr='{transparencyAttr}'");
            Console.WriteLine($"Debug: color='{color}', transparencyStr='{transparencyStr}'");

            if (!string.IsNullOrEmpty(transparencyStr) && int.TryParse(transparencyStr, out int transparencyValue))
            {
                if (transparencyValue < 0)
                    transparency = 0;
                else if (transparencyValue > 100)
                    transparency = 100;
                else
                    transparency = transparencyValue;
            }

            Console.WriteLine($"Debug: Final transparency={transparency}");
            return new ColorA(color, transparency);
        }
    }

    public enum ChartType
    {
        Bar,
        Column,
        Line,
        Area,
        Pie,
        Doughnut,
        Scatter,
        Bubble,
        Radar,
    }

    enum MarkerType
    {
        Circle,
        Square,
        Diamond,
        Triangle,
        X,
    }
    enum LegendPosition
    {
        Top,
        Bottom,
        Left,
        Right,
        TopRight,
    }
    enum DataLabelPosition
    {
        BestFit,
        Bottom,
        Center,
        InsideBase,
        InsideEnd,
        Left,
        OutsideEnd,
        Right,
        Top,
    }
    enum UnitType
    {
        Px,
        Cm,
        Pt,
        In,
    }
    enum BlanksDisplayedAs
    {
        Span,
        Gap,
        Zero
    }
    enum GroupingType
    {
        Standard,
        Clustered,
        Stacked,
        PercentStacked
    }
    enum LayoutMode
    {
        Edge,
        Factor,
    }
    enum LayoutTarget
    {
        Inner,
        Outer,
    }
    enum ThreeDShape
    {
        Cone,
        ConeToMax,
        Box,
        Cylinder,
        Pyramid,
        PyramidToMaximum,
    }
    enum PatternFillPreset
    {
        Percent5,
        Percent10,
        Percent20,
        Percent25,
        Percent30,
        Percent40,
        Percent50,
        Percent60,
        Percent70,
        Percent75,
        Percent80,
        Percent90,
        Horizontal,
        Vertical,
        LightHorizontal,
        LightVertical,
        DarkHorizontal,
        DarkVertical,
        NarrowHorizontal,
        NarrowVertical,
        DashedHorizontal,
        DashedVertical,
        Cross,
        DownwardDiagonal,
        UpwardDiagonal,
        LightDownwardDiagonal,
        LightUpwardDiagonal,
        DarkDownwardDiagonal,
        DarkUpwardDiagonal,
        WideDownwardDiagonal,
        WideUpwardDiagonal,
        DashedDownwardDiagonal,
        DashedUpwardDiagonal,
        DiagonalCross,
        SmallCheck,
        LargeCheck,
        SmallGrid,
        LargeGrid,
        DotGrid,
        SmallConfetti,
        LargeConfetti,
        HorizontalBrick,
        DiagonalBrick,
        SolidDiamond,
        OpenDiamond,
        DottedDiamond,
        Plaid,
        Sphere,
        Weave,
        Divot,
        Shingle,
        Wave,
        Trellis,
        ZigZag,
    }
    enum FillStyle
    {
        Auto, // (Varsayılan) öncelik: Gradient > Pattern > Solid
        Solid,
        Pattern,
        Gradient,
        None      // Hiç dolgu yok (NoFill)
    }
    enum ScatterStyle
    {
        Line,
        LineMarker,
        Marker,
        Smooth,
        SmoothMarker,
    }
    enum RadarStyle
    {
        Standard,   // sadece çizgi
        Marker,   // çizgi + marker
        Filled,   // dolu alan (area gibi)
    }
    enum ShadowType
    {
        Inner,
        Outer,
        Preset,
    }
    enum ShadowPreset
    {
        TopLeftDropShadow,
        TopRightDropShadow,
        BackLeftPerspectiveShadow,
        BackRightPerspectiveShadow,
        BottomLeftDropShadow,
        BottomRightDropShadow,
        FrontLeftPerspectiveShadow,
        FrontRightPerspectiveShadow,
        TopLeftSmallDropShadow,
        TopLeftLargeDropShadow,
        BackLeftLongPerspectiveShadow,
        BackRightLongPerspectiveShadow,
        TopLeftDoubleDropShadow,
        BottomRightSmallDropShadow,
        FrontLeftLongPerspectiveShadow,
        FrontRightLongPerspectiveShadow,
        ThreeDimensionalOuterBoxShadow,
        ThreeDimensionalInnerBoxShadow,
        BackCenterPerspectiveShadow,
        FrontBottomShadow,
    }
    enum BevelPreset
    {
        RelaxedInset,
        Circle,
        Slope,
        Cross,
        Angle,
        SoftRound,
        Convex,
        CoolSlant,
        Divot,
        Riblet,
        HardEdge,
        ArtDeco,
    }
    enum MaterialPreset
    {
        LegacyMatte,
        LegacyPlastic,
        LegacyMetal,
        LegacyWireframe,
        Matte,
        Plastic,
        Metal,
        WarmMatte,
        TranslucentPowder,
        Powder,
        DarkEdge,
        SoftEdge,
        Clear,
        Flat,
        SoftMetal,
    }
    enum LightingPreset
    {
        LegacyFlat1,
        LegacyFlat2,
        LegacyFlat3,
        LegacyFlat4,
        LegacyNormal1,
        LegacyNormal2,
        LegacyNormal3,
        LegacyNormal4,
        LegacyHarsh1,
        LegacyHarsh2,
        LegacyHarsh3,
        LegacyHarsh4,
        ThreePoints,
        Balanced,
        Soft,
        Harsh,
        Flood,
        Contrasting,
        Morning,
        Sunrise,
        Sunset,
        Chilly,
        Freezing,
        Flat,
        TwoPoints,
        Glow,
        BrightRoom,
    }
    enum LightingDirection
    {
        TopLeft,
        Top,
        TopRight,
        Left,
        Right,
        BottomLeft,
        Bottom,
        BottomRight,
    }
    enum DashPreset
    {
        Solid,
        Dot,
        Dash,
        LargeDash,
        DashDot,
        LargeDashDot,
        LargeDashDotDot,
        SystemDash,
        SystemDot,
        SystemDashDot,
        SystemDashDotDot,
    }
    enum CompoundLinePreset
    {
        Single,
        Double,
        ThickThin,
        ThinThick,
        Triple,
    }
    enum LineCapPreset
    {
        Round,
        Square,
        Flat,
    }
    enum LineJoinPreset
    {
        Round,
        Bevel,
        Miter,
    }
    enum LineEndPreset
    {
        None,
        Triangle,
        Stealth,
        Diamond,
        Oval,
        Arrow,
    }
    enum LineEndWidthPreset
    {
        Small,
        Medium,
        Large,
    }
    enum LineEndLengthPreset
    {
        Small,
        Medium,
        Large,
    }
    enum TextUnderlineStyle
    {
        None,
        Single,
        Double,
    }
    enum TextStrikeStyle
    {
        None,
        Single,
        Double,
    }
    enum TextCapsStyle
    {
        None,
        Small,
        All,
    }
    enum HorizontalAlign
    {
        Left,
        Center,
        Right,
        Justify,
    }
    enum VerticalAlign
    {
        Top,
        Center,
        Bottom,
    }
    enum TextWrapPreset
    {
        None,
        Square,
    }



}


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
            public int? GapWidth { get; set; }   // Bar, Column only
            public int? Overlap { get; set; }   // Bar, Column only
            public ScatterStyle ScatterStyle { get; set; }
            public RadarStyle RadarStyle { get; set; }
            public bool? AutoTitle { get; set; } //default: false
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
            public BlanksDisplayedAs? BlanksDisplayedAs { get; set; }
            public ThreeDViewDefinition ThreeDView { get; set; }
        }

        private static ChartDefinition GetChartDefinitionFromXml(XElement chartNode)
        {
            ChartDefinition chartDefinition = new ChartDefinition();

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

            if (XElementAttributeGetter.AsInt32(chartNode, "overlap", out int overlapValue))
                chartDefinition.Overlap = overlapValue;

            if (XElementAttributeGetter.AsInt32(chartNode, "gap-width", out int gapWidthValue))
                chartDefinition.GapWidth = gapWidthValue;




            return chartDefinition;
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
                    return true;    // Parse başarılı
                result = defaultValue;  // Parse başarısızsa default değeri ata
                return false;       // Parse başarısız
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
                ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
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
            addValueAxis(chartDefinition, chartNode, plotArea, categoryAxisId, valueAxisId, valueAxisPosition );

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

            chart.Append(new DisplayBlanksAs() { Val = mapBlanksDisplayedAs(chartDefinition.BlanksDisplayedAs ?? BlanksDisplayedAs.Gap) });
            chart.Append(new AutoTitleDeleted() { Val = (!chartDefinition.AutoTitle ?? false) });

            chart.Append(plotArea);

            return chart;
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

            chart.Append(new DisplayBlanksAs() { Val = mapBlanksDisplayedAs(chartDefinition.BlanksDisplayedAs ?? BlanksDisplayedAs.Gap) });
            chart.Append(new AutoTitleDeleted() { Val = (!chartDefinition.AutoTitle ?? false) });

            chart.Append(plotArea);

            return chart;
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

            chart.Append(new DisplayBlanksAs() { Val = mapBlanksDisplayedAs(chartDefinition.BlanksDisplayedAs ?? BlanksDisplayedAs.Gap) });
            chart.Append(new AutoTitleDeleted() { Val = (!chartDefinition.AutoTitle ?? false) });

            chart.Append(plotArea);

            return chart;
        }
        private static C.Chart getPieChart(ChartDefinition chartDefinition, XElement chartNode)
        {

            C.PlotArea plotArea = new C.PlotArea();
            plotArea.Append(new Layout());

            PieChart pieChart = new(new VaryColors() { Val = chartDefinition.VaryColors ?? false });

            uint seriesIndex = 0;

            foreach (XElement seriesNode in chartNode.Elements("series"))
            {
                PieChartSeries series  = new PieChartSeries(
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

            plotArea.Append(pieChart);

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

            chart.Append(new DisplayBlanksAs() { Val = mapBlanksDisplayedAs(chartDefinition.BlanksDisplayedAs ?? BlanksDisplayedAs.Gap) });
            chart.Append(new AutoTitleDeleted() { Val = (!chartDefinition.AutoTitle ?? false) });

            chart.Append(plotArea);

            return chart;
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

            doughnutChart.Append(new C.HoleSize() { Val = (ByteValue)(byte)50 });

            plotArea.Append(doughnutChart);

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

            chart.Append(new DisplayBlanksAs() { Val = mapBlanksDisplayedAs(chartDefinition.BlanksDisplayedAs ?? BlanksDisplayedAs.Gap) });
            chart.Append(new AutoTitleDeleted() { Val = (!chartDefinition.AutoTitle ?? false) });

            chart.Append(plotArea);

            return chart;
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

            chart.Append(new DisplayBlanksAs() { Val = mapBlanksDisplayedAs(chartDefinition.BlanksDisplayedAs ?? BlanksDisplayedAs.Gap) });
            chart.Append(new AutoTitleDeleted() { Val = (!chartDefinition.AutoTitle ?? false) });

            chart.Append(plotArea);

            return chart;
        }
        private static C.Chart getBubbleChart(ChartDefinition chartDefinition, XElement chartNode)
        {
            uint categoryAxisId = getSafeId();
            uint valueAxisId = getSafeId();

            C.PlotArea plotArea = new();
            plotArea.Append(new Layout());

            BubbleChart bubbleChart = new(
                new VaryColors() { Val =  chartDefinition.VaryColors ?? false }
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

            bubbleChart.Append(new C.BubbleScale() { Val = (UInt32Value)100 });
            bubbleChart.Append(new C.ShowNegativeBubbles() { Val = true });

            addAxisIds(bubbleChart, categoryAxisId, valueAxisId);

            plotArea.Append(bubbleChart);

            addValueAxis(chartDefinition, chartNode, plotArea, valueAxisId, categoryAxisId, position: AxisPositionValues.Bottom);
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

            chart.Append(new DisplayBlanksAs() { Val = mapBlanksDisplayedAs(chartDefinition.BlanksDisplayedAs ?? BlanksDisplayedAs.Gap) });
            chart.Append(new AutoTitleDeleted() { Val = (!chartDefinition.AutoTitle ?? false) });

            chart.Append(plotArea);

            return chart;

        }
        private static C.Chart getRadarChart(ChartDefinition chartDefinition, XElement chartNode)
        {
            uint categoryAxisId = getSafeId();
            uint valueAxisId = getSafeId();

            C.PlotArea plotArea = new();
            plotArea.Append(new Layout());

            RadarChart radarChart = new(
                new C.RadarStyle() { Val = RadarStyleValues.Standard },
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

            chart.Append(new DisplayBlanksAs() { Val = mapBlanksDisplayedAs(chartDefinition.BlanksDisplayedAs ?? BlanksDisplayedAs.Gap) });
            chart.Append(new AutoTitleDeleted() { Val = (!chartDefinition.AutoTitle ?? false) });

            chart.Append(plotArea);

            return chart;

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
                new AxisPosition() { Val = position },
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
                new AxisPosition() { Val = position },
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

            plotArea.Append(catAxis);
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

    enum ThreeDShape
    {
        Cone,
        ConeToMax,
        Box,
        Cylinder,
        Pyramid,
        PyramidToMaximum,
    }
    enum BlanksDisplayedAs
    {
        Span,
        Gap,
        Zero
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

    }


}


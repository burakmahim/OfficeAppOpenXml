using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2010.Excel;
using System.Globalization;
using System.Xml;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using System;
using DocumentFormat.OpenXml.Bibliography;
using System.Security.Cryptography;

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
            public int? GapWidth { get; set; }   // Bar, Column only    0 500
            public int? Overlap { get; set; }   // Bar, Column only     -100 100
            public ScatterStyle ScatterStyle { get; set; }
            public RadarStyle RadarStyle { get; set; }
            public bool? AutoTitle { get; set; } //default: false
            public bool? PlotVisibleOnly { get; set; }
            public bool? RoundedCorners { get; set; }
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

            Nullable<ChartType> chartType = XElementAttributeGetter.AsEnum<ChartType>(chartNode, "type");
            if (chartType.HasValue)
                chartDefinition.Type = chartType.Value;

            Nullable<RadarStyle> radarStyle = XElementAttributeGetter.AsEnum<RadarStyle>(chartNode, "radarStyle");
            if (radarStyle.HasValue)
                chartDefinition.RadarStyle = radarStyle.Value;

            if (XElementAttributeGetter.AsInt32(chartNode, "gap_width", out int gapWidth))
            {
                chartDefinition.GapWidth = gapWidth;
            }

            if (XElementAttributeGetter.AsInt32(chartNode, "overlap", out int overlap))
            {
                chartDefinition.Overlap = overlap;
            }

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

            public static bool AsBool(XElement element, string attributeName, out bool result)
            {
                result = false;
                string value = AsString(element, attributeName);
                return bool.TryParse(value, out result);
            }

            public static bool AsInt32(XElement element, string attributeName, out int result)
            {
                result = 0;
                string value = AsString(element, attributeName);
                return int.TryParse(value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out result);
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
        }
        public static void AddChart(WorksheetPart worksheetPart, XElement chartNode)
        {
            ChartDefinition chartDefinition = GetChartDefinitionFromXml(chartNode);

            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            if (drawingsPart.WorksheetDrawing == null)
                drawingsPart.WorksheetDrawing = new WorksheetDrawing();

            ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
            string relId = drawingsPart.GetIdOfPart(chartPart);

            ChartSpace chartSpace = new ChartSpace();
            chartSpace.Append(new EditingLanguage() { Val = "tr-TR" });

            Chart? chart = null;

            switch (chartDefinition.Type)
            {
                case ChartType.Bar:
                    chart = getBarChart(chartNode, chartDefinition, BarDirectionValues.Bar);
                    break;
                case ChartType.Column:
                    chart = getBarChart(chartNode, chartDefinition, BarDirectionValues.Column);
                    break;
                case ChartType.Line:
                    chart = getLineChart(chartNode);
                    break;
                case ChartType.Area:
                    chart = getAreaChart(chartNode);
                    break;
                case ChartType.Pie:
                    chart = getPieChart(chartNode);
                    break;
                case ChartType.Doughnut:
                    chart = getDoughnutChart(chartNode);
                    break;
                case ChartType.Scatter:
                    chart = getScatterChart(chartNode);
                    break;
                case ChartType.Bubble:
                    chart = getBubbleChart(chartNode);
                    break;
                case ChartType.Radar:
                    chart = getRadarChart(chartNode
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

        private static Chart getBarChart(XElement chartNode, ChartDefinition chartDefinition, BarDirectionValues direction)
        {
            uint categoryAxisId = getSafeId();
            uint valueAxisId = getSafeId();

            PlotArea plotArea = new PlotArea();
            plotArea.Append(new Layout());

            BarChart barChart = new BarChart()
            {
                BarDirection = new BarDirection() { Val = direction },
                BarGrouping = new BarGrouping() { Val = BarGroupingValues.Clustered },
                VaryColors = new VaryColors() { Val = false }
            };

            addAxisIds(barChart, categoryAxisId, valueAxisId);

            uint seriesIndex = 0;

            foreach (XElement seriesNode in chartNode.Elements("series"))
            {
                BarChartSeries series = new BarChartSeries(
                    new C.Index() { Val = seriesIndex },
                    new Order() { Val = seriesIndex },
                    new C.SeriesText(new NumericValue(seriesNode.Attribute("name")?.Value ?? ""))
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
                    
                    stringLiteral.Append(new StringPoint() { Index = (uint)i, NumericValue = new NumericValue(categoryName) });
                    numberLiteral.Append(new NumericPoint() { Index = (uint)i, NumericValue = new NumericValue(valueStr) });
                }

                catAxisData.Append(stringLiteral);
                values.Append(numberLiteral);

                series.Append(catAxisData);
                series.Append(values);

                barChart.Append(series);
                seriesIndex++;
            }

            if (chartDefinition.Overlap.HasValue)
                barChart.Append(new Overlap() { Val = new SByteValue((sbyte)chartDefinition.Overlap.Value) });

            if (chartDefinition.GapWidth.HasValue)
                barChart.Append(new GapWidth() { Val = new UInt16Value((ushort)chartDefinition.GapWidth.Value) });

            plotArea.Append(barChart);

            AxisPositionValues categoryAxisPosition = direction == BarDirectionValues.Bar ? AxisPositionValues.Left : AxisPositionValues.Bottom;
            AxisPositionValues valueAxisPosition = direction == BarDirectionValues.Bar ? AxisPositionValues.Bottom : AxisPositionValues.Left;

            addCategoryAxis(chartNode, plotArea, categoryAxisId, valueAxisId);
            addValueAxis(chartNode, plotArea, categoryAxisId, valueAxisId);

            Chart chart = new Chart();

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

            return chart;
        }
        private static Chart getLineChart(XElement chartNode)
        {
            uint categoryAxisId = getSafeId();
            uint valueAxisId = getSafeId();

            PlotArea plotArea = new PlotArea();
            plotArea.Append(new Layout());

            LineChart lineChart = new LineChart()
            {
                Grouping = new Grouping() { Val = GroupingValues.Standard },
                VaryColors = new VaryColors() { Val = false }
            };

            addAxisIds(lineChart, categoryAxisId, valueAxisId);

            uint seriesIndex = 0;

            foreach (XElement seriesNode in chartNode.Elements("series"))
            {
                LineChartSeries series = new LineChartSeries(
                    new C.Index() { Val = seriesIndex },
                    new Order() { Val = seriesIndex },
                    new C.SeriesText(new NumericValue(seriesNode.Attribute("name")?.Value ?? ""))
                );

                List<XElement> points = seriesNode.Elements("point").ToList();
                uint pointCount = (uint)points.Count;

                CategoryAxisData catAxisData = new CategoryAxisData();
                C.Values values = new C.Values();

                StringLiteral stringLiteral = new StringLiteral();
                NumberLiteral numberLiteral = new NumberLiteral();

                stringLiteral.Append(new PointCount() { Val = pointCount });
                //numberLiteral.Append(new PointCount() { Val = pointCount });

                for (int i = 0; i < pointCount; i++)
                {
                    string categoryName = points[i].Attribute("category")?.Value ?? $"Kategori {i + 1}";
                    string valueStr = points[i].Attribute("value")?.Value ?? "0";

                    stringLiteral.Append(new StringPoint() 
                    { 
                        Index = (uint)i, 
                        NumericValue = new NumericValue(categoryName) 
                    });

                    numberLiteral.Append(new NumericPoint() 
                    { 
                        Index = (uint)i, 
                        NumericValue = new NumericValue(valueStr) 
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

            addCategoryAxis(chartNode, plotArea, categoryAxisId, valueAxisId);
            addValueAxis(chartNode, plotArea, categoryAxisId, valueAxisId);

            Chart chart = new Chart();

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

            return chart;
        }
        private static Chart getAreaChart(XElement chartNode)
        {
            uint categoryAxisId = getSafeId();
            uint valueAxisId = getSafeId();

            PlotArea plotArea = new PlotArea();
            plotArea.Append(new Layout());

            AreaChart areaChart = new()
            {
                Grouping = new Grouping() { Val = GroupingValues.Standard },
                VaryColors = new VaryColors() { Val = false }
            };

            addAxisIds(areaChart, categoryAxisId, valueAxisId);

            uint seriesIndex = 0;

            foreach (XElement seriesNode in chartNode.Elements("series"))
            {
                AreaChartSeries series = new AreaChartSeries(
                    new C.Index() { Val = seriesIndex },
                    new Order() { Val = seriesIndex },
                    new C.SeriesText(new NumericValue(seriesNode.Attribute("name")?.Value ?? ""))
                );

                List<XElement> points = seriesNode.Elements("point").ToList();
                uint pointCount = (uint)points.Count;

                CategoryAxisData catAxisData = new CategoryAxisData();
                C.Values values = new C.Values();

                StringLiteral stringLiteral = new StringLiteral();
                NumberLiteral numberLiteral = new NumberLiteral();

                stringLiteral.Append(new PointCount() { Val = pointCount });
                //numberLiteral.Append(new PointCount() { Val = pointCount });

                for (int i = 0; i < pointCount; i++)
                {
                    string categoryName = points[i].Attribute("category")?.Value ?? $"Kategori {i + 1}";
                    string valueStr = points[i].Attribute("value")?.Value ?? "0";

                    stringLiteral.Append(new StringPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new NumericValue(categoryName)
                    });

                    numberLiteral.Append(new NumericPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new NumericValue(valueStr)
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

            addCategoryAxis(chartNode, plotArea, categoryAxisId, valueAxisId);
            addValueAxis(chartNode, plotArea, categoryAxisId, valueAxisId);

            Chart chart = new Chart();

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

            return chart;
        }
        private static Chart getPieChart(XElement chartNode)
        {

            PlotArea plotArea = new PlotArea();
            plotArea.Append(new Layout());

            PieChart pieChart = new(new VaryColors() { Val = true });

            uint seriesIndex = 0;

            foreach (XElement seriesNode in chartNode.Elements("series"))
            {
                PieChartSeries series  = new PieChartSeries(
                    new C.Index() { Val = seriesIndex },
                    new Order() { Val = seriesIndex },
                    new C.SeriesText(new NumericValue(seriesNode.Attribute("name")?.Value ?? ""))
                );

                List<XElement> points = seriesNode.Elements("point").ToList();
                uint pointCount = (uint)points.Count;

                CategoryAxisData catAxisData = new CategoryAxisData();
                C.Values values = new C.Values();

                StringLiteral stringLiteral = new StringLiteral();
                NumberLiteral numberLiteral = new NumberLiteral();

                stringLiteral.Append(new PointCount() { Val = pointCount });
                //numberLiteral.Append(new PointCount() { Val = pointCount });

                for (int i = 0; i < pointCount; i++)
                {
                    string categoryName = points[i].Attribute("category")?.Value ?? $"Kategori {i + 1}";
                    string valueStr = points[i].Attribute("value")?.Value ?? "0";

                    stringLiteral.Append(new StringPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new NumericValue(categoryName)
                    });

                    numberLiteral.Append(new NumericPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new NumericValue(valueStr)
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

            Chart chart = new Chart();

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

            return chart;
        }
        private static Chart getDoughnutChart(XElement chartNode)
        {

            PlotArea plotArea = new PlotArea();
            plotArea.Append(new Layout());

            DoughnutChart doughnutChart = new(new VaryColors() { Val = true });

            uint seriesIndex = 0;

            foreach (XElement seriesNode in chartNode.Elements("series"))
            {
                PieChartSeries series = new PieChartSeries(
                    new C.Index() { Val = seriesIndex },
                    new Order() { Val = seriesIndex },
                    new C.SeriesText(new NumericValue(seriesNode.Attribute("name")?.Value ?? ""))
                );

                List<XElement> points = seriesNode.Elements("point").ToList();
                uint pointCount = (uint)points.Count;

                CategoryAxisData catAxisData = new CategoryAxisData();
                C.Values values = new C.Values();

                StringLiteral stringLiteral = new StringLiteral();
                NumberLiteral numberLiteral = new NumberLiteral();

                stringLiteral.Append(new PointCount() { Val = pointCount });
                //numberLiteral.Append(new PointCount() { Val = pointCount });

                for (int i = 0; i < pointCount; i++)
                {
                    string categoryName = points[i].Attribute("category")?.Value ?? $"Kategori {i + 1}";
                    string valueStr = points[i].Attribute("value")?.Value ?? "0";

                    stringLiteral.Append(new StringPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new NumericValue(categoryName)
                    });

                    numberLiteral.Append(new NumericPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new NumericValue(valueStr)
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

            Chart chart = new Chart();

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

            return chart;
        }
        private static Chart getScatterChart(XElement chartNode)
        {
            uint categoryAxisId = getSafeId();
            uint valueAxisId = getSafeId();

            PlotArea plotArea = new PlotArea();
            plotArea.Append(new Layout());

            ScatterChart scatterChart = new(
                new C.ScatterStyle() { Val = ScatterStyleValues.Marker },
                new C.VaryColors() { Val = false }
            );

            uint seriesIndex = 0;

            foreach (XElement seriesNode in chartNode.Elements("series"))
            {
                ScatterChartSeries series = new ScatterChartSeries(
                    new C.Index() { Val = seriesIndex },
                    new Order() { Val = seriesIndex },
                    new C.SeriesText(new NumericValue(seriesNode.Attribute("name")?.Value ?? ""))
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

            addValueAxis(chartNode, plotArea, valueAxisId, categoryAxisId, position: AxisPositionValues.Bottom);
            addValueAxis(chartNode, plotArea, categoryAxisId, valueAxisId, position: AxisPositionValues.Left);

            Chart chart = new Chart();

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

            return chart;
        }
        private static Chart getBubbleChart(XElement chartNode)
        {
            uint categoryAxisId = getSafeId();
            uint valueAxisId = getSafeId();

            PlotArea plotArea = new();
            plotArea.Append(new Layout());

            BubbleChart bubbleChart = new(
                new VaryColors() { Val =  false }
            );

            uint seriesIndex = 0;

            foreach (XElement seriesNode in chartNode.Elements("series"))
            {
                BubbleChartSeries series = new BubbleChartSeries(
                    new C.Index() { Val = seriesIndex },
                    new Order() { Val = seriesIndex },
                    new C.SeriesText(new NumericValue(seriesNode.Attribute("name")?.Value ?? ""))
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

            addValueAxis(chartNode, plotArea, valueAxisId, categoryAxisId, position: AxisPositionValues.Bottom);
            addValueAxis(chartNode, plotArea, categoryAxisId, valueAxisId, position: AxisPositionValues.Left);

            Chart chart = new Chart();

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

            return chart;

        }
        private static Chart getRadarChart(XElement chartNode)
        {
            uint categoryAxisId = getSafeId();
            uint valueAxisId = getSafeId();

            PlotArea plotArea = new();
            plotArea.Append(new Layout());

            RadarChart radarChart = new(
                new C.RadarStyle() { Val = RadarStyleValues.Standard },
                new C.VaryColors() { Val = false }
            );

            uint seriesIndex = 0;

            foreach (XElement seriesNode in chartNode.Elements("series"))
            {
                RadarChartSeries series = new RadarChartSeries(
                    new C.Index() { Val = seriesIndex },
                    new Order() { Val = seriesIndex },
                    new C.SeriesText(new NumericValue(seriesNode.Attribute("name")?.Value ?? ""))
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
                        NumericValue = new NumericValue(categoryName)
                    });
                    numberLiteral.Append(new NumericPoint()
                    {
                        Index = (uint)i,
                        NumericValue = new NumericValue(valueStr)
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

            addCategoryAxis(chartNode, plotArea, categoryAxisId, valueAxisId, position: AxisPositionValues.Bottom);
            addValueAxis(chartNode, plotArea, categoryAxisId, valueAxisId, position: AxisPositionValues.Left);

            Chart chart = new Chart();

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
        private static void addValueAxis(XElement chartNode, PlotArea plotArea, uint categoryAxisId, uint valueAxisId, AxisPositionValues? position = null) 
        {
            AxisPositionValues axisPos = position ?? AxisPositionValues.Left;

            ValueAxis valAx = new ValueAxis(
                new AxisId() { Val = valueAxisId },
                new Delete() { Val = false },
                new Scaling(new Orientation() { Val = C.OrientationValues.MinMax }),
                new AxisPosition() { Val = position },
                new C.MajorGridlines(),
                new TickLabelPosition() { Val = TickLabelPositionValues.NextTo },
                new CrossingAxis() { Val = categoryAxisId },
                new Crosses() { Val = CrossesValues.AutoZero },
                new CrossBetween() { Val = CrossBetweenValues.Between }
            );

            plotArea.Append(valAx);
        }
        private static void addCategoryAxis(XElement chartNode, PlotArea plotArea, uint categoryAxisId, uint valueAxisId, AxisPositionValues? position = null)
        {
            AxisPositionValues axisPos = position ?? AxisPositionValues.Bottom;

            CategoryAxis catAxis = new CategoryAxis(
                new AxisId() { Val = categoryAxisId },
                new Delete() { Val = false },
                new Scaling(new Orientation() { Val = C.OrientationValues.MinMax }),
                new AxisPosition() { Val = position },
                new TickLabelPosition() { Val = TickLabelPositionValues.NextTo },
                new CrossingAxis() { Val = valueAxisId },
                new Crosses() { Val = CrossesValues.AutoZero },
                new AutoLabeled() { Val = true },
                new LabelAlignment() { Val = LabelAlignmentValues.Center }
            );

            plotArea.Append(catAxis);
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

}


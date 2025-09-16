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
using System.Text.RegularExpressions;
using System.Xml;
using DocumentFormat.OpenXml.VariantTypes;
using OfficeAppOpenXmlLibrary.PowerPointOpenXmlComponents;
using System;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

namespace OfficeAppOpenXmlLibrary.Common
{
    public static class OfficeDocument   
    {
        private             void                        addChart                (ChartDefinition chartDefinition)                                                   
        {
            if (chartDefinition is null || chartDefinition.Series.Count == 0)
                return;

            Chart chart = null;

            switch (chartDefinition.Type)
            {
                case ChartType.Bar     :
                    chart = getBarChart(chartDefinition, BarDirectionValues.Bar);
                    break;
                case ChartType.Column  :
                    chart = getBarChart(chartDefinition, BarDirectionValues.Column);
                    break;
                case ChartType.Line    :
                    chart = getLineChart(chartDefinition);
                    break;
                case ChartType.Area    :
                    chart = getAreaChart(chartDefinition);
                    break;
                case ChartType.Pie     :
                    chart = getPieChart(chartDefinition);
                    break;
                case ChartType.Doughnut:
                    chart = getDoughnutChart(chartDefinition);
                    break;
                case ChartType.Scatter :
                    chart = getScatterChart(chartDefinition);
                    break;
                case ChartType.Bubble  :
                    chart = getBubbleChart(chartDefinition);
                    break;
                case ChartType.Radar   :
                    chart = getRadarChart(chartDefinition);
                    break;
                default                :
                    break;
            }

            if (chart is null)
                return;

            ChartPart  chartPart  = assets.WordDocumentCapsule.WordProcessingDocument.MainDocumentPart.AddNewPart<ChartPart>();
            string     relId      = assets.WordDocumentCapsule.WordProcessingDocument.MainDocumentPart.GetIdOfPart(chartPart);
            ChartSpace chartSpace = new ();

            if (chartDefinition.ChartAreaFormat is not null)
            {
                applyFormat(chartSpace, chartDefinition.ChartAreaFormat);
            }

            if (chartDefinition.TextFormat is not null)
            {
                C.TextProperties textProperties = new C.TextProperties();

                if (applyTextFormat(textProperties, chartDefinition.TextFormat))
                    chartSpace.Append(textProperties);
            }

            chartSpace.Append(new RoundedCorners() { Val = chartDefinition.RoundedCorners ?? false });

            chartSpace.Append(chart);

            chartPart.ChartSpace = chartSpace;

            Paragraph paragraph;

            if (chartDefinition.CreateNewParagraph)
            {
                paragraph = assets.ContainerDefinitionStack.Definition?.CreateParagraph();
            }
            else
            {
                paragraph = assets.ContainerDefinitionStack.Definition?.ActiveParagraph ?? assets.ContainerDefinitionStack.Definition?.CreateParagraph();
            }

            if (paragraph is null)
                return;

            Run run = new Run();

            run.Append(new Drawing(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = chartDefinition.Width.ToEmu(), Cy = chartDefinition.Height.ToEmu() },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties() { Id = (UInt32Value)1U, Name = "Chart " + Guid.NewGuid() },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                        new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks() { NoChangeAspect = true }),
                    new DocumentFormat.OpenXml.Drawing.Graphic(
                        new DocumentFormat.OpenXml.Drawing.GraphicData(
                            new DocumentFormat.OpenXml.Drawing.Charts.ChartReference() { Id = relId }
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" })
                )
                {
                    DistanceFromTop    = (UInt32Value)0U,
                    DistanceFromBottom = (UInt32Value)0U,
                    DistanceFromLeft   = (UInt32Value)0U,
                    DistanceFromRight  = (UInt32Value)0U,
                }));

            paragraph.Append(run);
        }

        private     static  Chart                       getBarChart             (ChartDefinition chartDefinition, BarDirectionValues barDirection)                  
        {
            uint categoryAxisId = getSafeId();
            uint valueAxisId    = getSafeId();

            PlotArea plotArea = new ();
            AxisPositionValues categoryAxisPosition = barDirection == BarDirectionValues.Column ? AxisPositionValues.Bottom : AxisPositionValues.Left;
            AxisPositionValues valueAxisPosition    = barDirection == BarDirectionValues.Column ? AxisPositionValues.Left   : AxisPositionValues.Bottom;

            if (chartDefinition.ThreeDView is not null)
            {
                Bar3DChart bar3DChart = new Bar3DChart()
                {
                    BarDirection = new BarDirection() { Val = barDirection },
                    BarGrouping  = new BarGrouping () { Val = mapBarGrouping(chartDefinition.GroupingType ?? GroupingType.Clustered) },
                    VaryColors   = new VaryColors  () { Val = chartDefinition.VaryColors ?? false }
                };

                addAxisIds(bar3DChart, categoryAxisId, valueAxisId);

                applyBar3DOptions(bar3DChart, chartDefinition.ThreeDView);

                addBarSeries(chartDefinition, bar3DChart);

                addDataLabels(bar3DChart, chartDefinition);

                plotArea.Append(bar3DChart);
            }
            else
            {
                BarChart barChart = new BarChart()
                {
                    BarDirection = new BarDirection() { Val = barDirection },
                    BarGrouping  = new BarGrouping () { Val = mapBarGrouping(chartDefinition.GroupingType ?? GroupingType.Clustered) },
                    VaryColors   = new VaryColors  () { Val = chartDefinition.VaryColors ?? false }
                };

                addAxisIds(barChart, categoryAxisId, valueAxisId);

                addBarSeries(chartDefinition, barChart);

                if (chartDefinition.Overlap.HasValue)
                    barChart.Append(new Overlap() { Val = new SByteValue((sbyte)chartDefinition.Overlap.Value) });

                if (chartDefinition.GapWidth.HasValue)
                    barChart.Append(new GapWidth() { Val = new UInt16Value((ushort)chartDefinition.GapWidth.Value) });

                addDataLabels(barChart, chartDefinition);

                plotArea.Append(barChart);
            }

            addCategoryAxis (chartDefinition, plotArea, categoryAxisId, valueAxisId, categoryAxisPosition);
            addValueAxis    (chartDefinition, plotArea, categoryAxisId, valueAxisId, valueAxisPosition);
            addLayout       (chartDefinition.Layout, plotArea);
            //addWallsAndFloor(plotArea, chartDefinition.ThreeDViewDefinition);

            return getChartCommon(chartDefinition, plotArea);
        }
        private     static  Chart                       getLineChart            (ChartDefinition chartDefinition)                                                   
        {
            uint catId = getSafeId();
            uint valId = getSafeId();

            PlotArea plotArea = new ();

            bool use3D = chartDefinition.ThreeDView is not null;

            if (use3D)
            {
                Line3DChart line3DChart = new Line3DChart()
                {
                    Grouping   = new Grouping()  { Val = mapGrouping(chartDefinition.GroupingType ?? GroupingType.Standard) },
                    VaryColors = new VaryColors(){ Val = chartDefinition.VaryColors ?? false }
                };

                addAxisIds   (line3DChart, catId, valId);
                addLineSeries(chartDefinition, line3DChart);              
                addDataLabels(line3DChart, chartDefinition);

                plotArea.Append(line3DChart);

                addCategoryAxis (chartDefinition, plotArea, catId, valId);
                addValueAxis    (chartDefinition, plotArea, catId, valId);
                addLayout       (chartDefinition.Layout, plotArea);

                return getChartCommon(chartDefinition, plotArea);
            }
            else
            {
                LineChart lineChart = new()
                    {
                        Grouping   = new Grouping() { Val = mapGrouping(chartDefinition.GroupingType ?? GroupingType.Standard) },
                        VaryColors = new VaryColors() { Val = chartDefinition.VaryColors ?? false }
                    };

                addAxisIds(lineChart, catId, valId);

                addLineSeries(chartDefinition, lineChart);
                addDataLabels(lineChart, chartDefinition);

                plotArea.Append(lineChart);

                addCategoryAxis(chartDefinition, plotArea, catId, valId);
                addValueAxis   (chartDefinition, plotArea, catId, valId);
                addLayout      (chartDefinition.Layout, plotArea);

                return getChartCommon(chartDefinition, plotArea);
            }

        }
        private     static  Chart                       getAreaChart            (ChartDefinition chartDefinition)                                                   
        {
            uint catId = getSafeId();
            uint valId = getSafeId();

            PlotArea plotArea = new ();

            bool use3D = chartDefinition.ThreeDView is not null;

            if (use3D)
            {
                Area3DChart area3DChart = new ()
                {
                    Grouping   = new Grouping() { Val = mapGrouping(chartDefinition.GroupingType ?? GroupingType.Standard) },
                    VaryColors = new VaryColors() { Val = chartDefinition.VaryColors ?? false }
                };

                addAxisIds   (area3DChart, catId, valId);
                addAreaSeries(chartDefinition, area3DChart);
                addDataLabels(area3DChart, chartDefinition);

                plotArea.Append(area3DChart);

                addCategoryAxis (chartDefinition, plotArea, catId, valId);
                addValueAxis    (chartDefinition, plotArea, catId, valId);
                addLayout       (chartDefinition.Layout, plotArea);

                return getChartCommon(chartDefinition, plotArea);
            }
            else
            {
                AreaChart areaChart = new()
                    {
                        Grouping   = new Grouping() { Val = mapGrouping(chartDefinition.GroupingType ?? GroupingType.Standard) },
                        VaryColors = new VaryColors() { Val = chartDefinition.VaryColors ?? false }
                    };

                addAxisIds   (areaChart, catId, valId);
                addAreaSeries(chartDefinition, areaChart);
                addDataLabels(areaChart, chartDefinition);

                plotArea.Append(areaChart);

                addCategoryAxis(chartDefinition, plotArea, catId, valId);
                addValueAxis   (chartDefinition, plotArea, catId, valId);
                addLayout      (chartDefinition.Layout, plotArea);

                return getChartCommon(chartDefinition, plotArea);
            }
        }
        private     static  Chart                       getPieChart             (ChartDefinition chartDefinition)                                                   
        {
            PlotArea plotArea = new();

            bool use3D      = chartDefinition.ThreeDView is not null;

            if (use3D)
            {
                C.Pie3DChart pie3D = new (new VaryColors() { Val = chartDefinition.VaryColors ?? true });

                addPieSeries(chartDefinition, pie3D);
                addDataLabels         (pie3D, chartDefinition);

                if (chartDefinition.FirstSliceAngle.HasValue)
                    pie3D.Append(new C.FirstSliceAngle() { Val = (UInt16Value)(ushort)clamp(chartDefinition.FirstSliceAngle.Value, 0, 360) });

                plotArea.Append(pie3D);

                addLayout(chartDefinition.Layout, plotArea);

                // Pie3D'de view3D istemiyoruz
                return getChartCommon(chartDefinition, plotArea, allowView3D: false);
            }
            else
            {
                C.PieChart pieChart = new(new VaryColors() { Val = chartDefinition.VaryColors ?? true });

                addPieSeries(chartDefinition, pieChart);
                addDataLabels         (pieChart, chartDefinition);

                if (chartDefinition.FirstSliceAngle.HasValue)
                    pieChart.Append(new C.FirstSliceAngle() { Val = (UInt16Value)(ushort)clamp(chartDefinition.FirstSliceAngle.Value, 0, 360) });

                plotArea.Append(pieChart);

                addLayout(chartDefinition.Layout, plotArea);

                return getChartCommon(chartDefinition, plotArea);
            }
        }
        private     static  Chart                       getDoughnutChart        (ChartDefinition chartDefinition)                                                   
        {
            PlotArea plotArea = new();

            C.DoughnutChart doughnut = new (
                new VaryColors() { Val = chartDefinition.VaryColors ?? true }
            );

            addPieSeries(chartDefinition, doughnut);
            addDataLabels         (doughnut, chartDefinition);

            if (chartDefinition.FirstSliceAngle.HasValue)
                doughnut.Append(new C.FirstSliceAngle() { Val = (UInt16Value)(ushort)clamp(chartDefinition.FirstSliceAngle.Value, 0, 360) });

            doughnut.Append(new C.HoleSize() { Val = (ByteValue)(byte)clamp(chartDefinition.DoughnutHoleSize ?? 50, 10, 90) });

            plotArea.Append(doughnut);

            addLayout(chartDefinition.Layout, plotArea);

            return getChartCommon(chartDefinition, plotArea);
        }
        private     static  Chart                       getBubbleChart          (ChartDefinition chartDefinition)                                                   
        {
            // İki value axis ID’si (X ve Y)
            uint xId = getSafeId();
            uint yId = getSafeId();

            PlotArea plotArea = new();

            C.BubbleChart bubbleChart = new (
                new VaryColors() { Val = chartDefinition.VaryColors ?? false }
            );

            addBubbleSeries (chartDefinition, bubbleChart);
            addDataLabels   (bubbleChart, chartDefinition);

            if (chartDefinition.BubbleScale.HasValue)
                bubbleChart.Append(new C.BubbleScale() { Val = (UInt32Value)(uint)clamp(chartDefinition.BubbleScale.Value, 0, 300) });

            bubbleChart.Append(new C.ShowNegativeBubbles() { Val = true });

            addAxisIds      (bubbleChart, xId, yId);

            plotArea.Append(bubbleChart);

            addValueAxis(chartDefinition, plotArea, yId, xId, position: AxisPositionValues.Bottom);
            addValueAxis(chartDefinition, plotArea, xId, yId);

            addLayout(chartDefinition.Layout, plotArea);

            return getChartCommon(chartDefinition, plotArea);
        }
        private     static  Chart                       getScatterChart         (ChartDefinition chartDefinition)                                                   
        {
            uint xId = getSafeId();
            uint yId = getSafeId();

            PlotArea plotArea = new ();
            C.ScatterStyleValues scatterStyle = mapScatterStyle(chartDefinition.ScatterStyle);

            C.ScatterChart scatter = new (
                new C.ScatterStyle() { Val = scatterStyle },
                new C.VaryColors()   { Val = chartDefinition.VaryColors ?? false }
            );

            addScatterSeries(chartDefinition, scatter, scatterStyle);
            addDataLabels   (scatter, chartDefinition);
            addAxisIds      (scatter, xId, yId);

            plotArea.Append(scatter);

            // X = Bottom, Y = Left (numerik çift eksen)
            addValueAxis(chartDefinition, plotArea, yId, xId, position: AxisPositionValues.Bottom); // X
            addValueAxis(chartDefinition, plotArea, xId, yId, position: AxisPositionValues.Left  ); // Y

            addLayout(chartDefinition.Layout, plotArea);

            return getChartCommon(chartDefinition, plotArea);
        }
        private     static  Chart                       getRadarChart           (ChartDefinition chartDefinition)                                                   
        {
            uint catId = getSafeId();
            uint valId = getSafeId();

            PlotArea plotArea = new ();

            C.RadarChart radar = new (
                new C.RadarStyle { Val = mapRadarStyle(chartDefinition.RadarStyle) },
                new C.VaryColors { Val = chartDefinition.VaryColors ?? false }
            );

            addRadarSeries(chartDefinition, radar);

            addAxisIds(radar, catId, valId);
            plotArea.Append(radar);

            addCategoryAxis(chartDefinition, plotArea, catId, valId, AxisPositionValues.Bottom);
            addValueAxis   (chartDefinition, plotArea, catId, valId, AxisPositionValues.Left  );

            addLayout(chartDefinition.Layout, plotArea);

            return getChartCommon(chartDefinition, plotArea);
        }
        private     static  Chart                       getChartCommon          (ChartDefinition chartDefinition, PlotArea plotArea, bool allowView3D = true)       
        {
            if (plotArea is null)
                return null;

           Chart chart = new ();

            addTitle(chartDefinition, chart);

            if (allowView3D && chartDefinition.ThreeDView is not null)
            {
                add3DView(chart, chartDefinition);

                //if (chartDefinition.ThreeDViewDefinition is not null)
                //    switch (chartDefinition.Type)
                //    {
                //        case ChartType.Bar   :
                //        case ChartType.Column:
                //        case ChartType.Line  :
                //        case ChartType.Area  :
                //            add3DFloorAndWalls(chart, chartDefinition.ThreeDViewDefinition);
                //            break;
                //        default:
                //            break;
                //    }
            }

            chart.Append(plotArea);

            chart.Append(new AutoTitleDeleted() { Val = !(chartDefinition.AutoTitle ?? false) });

            chart.Append(new DisplayBlanksAs() { Val = mapBlanksDisplayedAs(chartDefinition.BlanksDisplayedAs ?? BlanksDisplayedAs.Gap) });

            chart.Append(new PlotVisibleOnly() { Val = chartDefinition.PlotVisibleOnly ?? false });

            chart.Append(new ShowDataLabelsOverMaximum() { Val = chartDefinition.ShowDataLabelsOverMaximum ?? true });

            if (chartDefinition.Legend is LegendDefinition legendDefinition && legendDefinition.Show)
            {
                Legend legend = new Legend(
                    new C.LegendPosition() { Val = mapLegendPosition(chartDefinition.Legend.Position) },
                    new Layout(),
                    new Overlay() { Val = false }
                );

                if (legendDefinition.LegendEntryTextFormat is not null)
                {
                    C.TextProperties textProperties = new();

                    if (applyTextFormat(textProperties, legendDefinition.LegendEntryTextFormat))
                        legend.Append(textProperties);
                }

                int index = 0;

                foreach (SeriesDefinition series in chartDefinition.Series)
                {
                    if (series.Title?.TextFormat is TextFormatDefinition textFormatDefinition)
                    {
                        C.TextProperties textProperties = new();

                        C.LegendEntry legendEntry = new (
                            new C.Index() { Val = (uint)index },
                            new C.Delete() { Val = false },
                            textProperties
                        );

                        if (applyTextFormat(textProperties, textFormatDefinition))
                            legend.Append(legendEntry);
                    }

                    index++;
                }

                if (legendDefinition.BoxFormat is not null)
                {
                    applyFormat(legend, legendDefinition.BoxFormat);
                }

                chart.Append(legend);
            }

            if (chartDefinition.DataTable is DataTableDefinition dataTableDefinition && dataTableDefinition.Show)
            {
                C.DataTable dataTable = new (
                    new C.ShowHorizontalBorder() { Val = dataTableDefinition.ShowHorizontalBorder },
                    new C.ShowVerticalBorder  () { Val = dataTableDefinition.ShowVerticalBorder   },
                    new C.ShowOutlineBorder   () { Val = dataTableDefinition.ShowOutlineBorder    },
                    new C.ShowKeys            () { Val = dataTableDefinition.ShowLegendKey        }
                    );

                if (dataTableDefinition.BoxFormat is not null)
                    applyFormat(dataTable, dataTableDefinition.BoxFormat);   // c:spPr

                if (dataTableDefinition.TextFormat is not null)
                {
                    C.TextProperties txPr = new ();
                    if (applyTextFormat(txPr, dataTableDefinition.TextFormat))
                        dataTable.Append(txPr);                // c:txPr
                }

                plotArea.Append(dataTable);
            }

            if (chartDefinition.PlotAreaFormat is not null)
            {
                applyFormat(chart.PlotArea, chartDefinition.PlotAreaFormat);
            }

            return chart;
        }

        private     static  void                        addBarSeries            (ChartDefinition chartDefinition, OpenXmlCompositeElement barChart)                 
        {
            uint seriesIndex = 0;

            foreach (SeriesDefinition seriesDefinition in chartDefinition.Series)
            {
                BarChartSeries series = new BarChartSeries(
                    new Index() { Val = seriesIndex },
                    new Order() { Val = seriesIndex },
                    new SeriesText(new NumericValue() { Text = seriesDefinition.Title })
                );
                
                CategoryAxisData catAxisData = new CategoryAxisData();
                Values values = new Values();

                StringLiteral stringLiteral = new StringLiteral();
                NumberLiteral numberLiteral = new NumberLiteral();

                stringLiteral.Append(new PointCount() { Val = (uint)seriesDefinition.Points.Count });

                for (int i = 0; i < seriesDefinition.Points.Count; i++)
                {
                    stringLiteral.Append(new StringPoint() { Index = (uint)i, NumericValue = new NumericValue(seriesDefinition.Points[i].Category) });
                    numberLiteral.Append(new NumericPoint() { Index = (uint)i, NumericValue = new NumericValue(seriesDefinition.Points[i].Value?.ToString(CultureInfo.InvariantCulture)) });
                }

                catAxisData.Append(stringLiteral);
                values     .Append(numberLiteral);

                applyPointLevelStyling(seriesDefinition, series, useMarker: false);

                series.Append(catAxisData);
                series.Append(values);

                applySeriesFormat(chartDefinition, seriesDefinition, series);
                addDataLabels(series, seriesDefinition);

                barChart.Append(series);

                seriesIndex++;
            }
        }
        private     static  void                        addLineSeries           (ChartDefinition chartDefinition, OpenXmlCompositeElement lineChart)                
        {
            uint idx = 0;

            foreach (SeriesDefinition seriesDefinition in chartDefinition.Series)
            {
                LineChartSeries series = new (
                    new Index() { Val = idx },
                    new Order() { Val = idx },
                    new SeriesText(new NumericValue() { Text = seriesDefinition.Title })
                );

                buildCategoryAndValues(seriesDefinition, out CategoryAxisData cat, out Values vals);

                ensureMarker(series, chartDefinition, seriesDefinition);
                applyPointLevelStyling(seriesDefinition, series, useMarker: true);

                series.Append(cat);
                series.Append(vals);

                applySeriesFormat(chartDefinition, seriesDefinition, series, toOutline: true);
                addDataLabels(series, seriesDefinition);

                lineChart.Append(series);

                idx++;
            }
        }
        private     static  void                        addAreaSeries           (ChartDefinition chartDefinition, OpenXmlCompositeElement areaChart)                
        {
            uint idx = 0;

            foreach (SeriesDefinition seriesDefinition in chartDefinition.Series)
            {
                AreaChartSeries series = new (
                    new Index() { Val = idx },
                    new Order() { Val = idx },
                    new SeriesText(new NumericValue() { Text = seriesDefinition.Title })
                );

                buildCategoryAndValues(seriesDefinition, out CategoryAxisData cat, out Values vals);

                series.Append(cat);
                series.Append(vals);

                applySeriesFormat(chartDefinition, seriesDefinition, series);
                addDataLabels(series, seriesDefinition);

                areaChart.Append(series);

                idx++;
            }
        }
        private     static  void                        addPieSeries            (ChartDefinition chartDefinition, OpenXmlCompositeElement owner)                    
        {
            uint idx = 0;

            foreach (SeriesDefinition seriesDefinition in chartDefinition.Series)
            {
                C.PieChartSeries series = new (
                    new Index() { Val = idx },
                    new Order() { Val = idx },
                    new SeriesText(new NumericValue() { Text = seriesDefinition.Title })
                );

                buildCategoryAndValues(seriesDefinition, out CategoryAxisData cat, out Values vals);
                applyPointLevelStyling(seriesDefinition, series, useMarker: false);

                series.Append(cat);
                series.Append(vals);

                applySeriesFormat(chartDefinition, seriesDefinition, series);
                addDataLabels(series, seriesDefinition);

                owner.Append(series);

                idx++;
            }
        }
        private     static  void                        addBubbleSeries         (ChartDefinition chartDefinition, C.BubbleChart bubbleChart)                        
        {
            uint idx = 0;

            foreach (SeriesDefinition seriesDefinition in chartDefinition.Series)
            {
                C.BubbleChartSeries series = new (
                    new Index() { Val = idx },
                    new Order() { Val = idx },
                    new SeriesText(new NumericValue() { Text = seriesDefinition.Title })
                );

                bool bubble3DEnabled = chartDefinition.Bubble3D ?? false;

                if (bubble3DEnabled)
                    series.Append(new C.InvertIfNegative() { Val = true });

                NumberLiteral xNumLit = new ();
                xNumLit.Append(new PointCount() { Val = (uint)seriesDefinition.Points.Count });
                for (int i = 0; i < seriesDefinition.Points.Count; i++)
                    xNumLit.Append(new NumericPoint() { Index = (uint)i, NumericValue = new NumericValue(seriesDefinition.Points[i].X?.ToString(CultureInfo.InvariantCulture)) });

                NumberLiteral yNumLit = new ();
                yNumLit.Append(new PointCount() { Val = (uint)seriesDefinition.Points.Count });
                for (int i = 0; i < seriesDefinition.Points.Count; i++)
                    yNumLit.Append(new NumericPoint() { Index = (uint)i, NumericValue = new NumericValue(seriesDefinition.Points[i].Y?.ToString(CultureInfo.InvariantCulture)) });

                NumberLiteral szNumLit = new ();
                szNumLit.Append(new PointCount() { Val = (uint)seriesDefinition.Points.Count });
                for (int i = 0; i < seriesDefinition.Points.Count; i++)
                    szNumLit.Append(new NumericPoint() { Index = (uint)i, NumericValue = new NumericValue(seriesDefinition.Points[i].Size?.ToString(CultureInfo.InvariantCulture)) });

                applyPointLevelStyling(seriesDefinition, series, useMarker: false);

                series.Append(new C.XValues(xNumLit));
                series.Append(new C.YValues(yNumLit));
                series.Append(new C.BubbleSize(szNumLit));

                if (bubble3DEnabled)
                    series.Append(new C.Bubble3D() { Val = true });

                applySeriesFormat(chartDefinition, seriesDefinition, series);
                addDataLabels(series, seriesDefinition);

                bubbleChart.Append(series);

                idx++;
            }
        }
        private     static  void                        addScatterSeries        (ChartDefinition chartDefinition, C.ScatterChart scatterChart, C.ScatterStyleValues scatterStyle)   
        {
            bool wantsLine   = scatterStyle is C.ScatterStyleValues.Line
                                            or C.ScatterStyleValues.LineMarker
                                            or C.ScatterStyleValues.Smooth
                                            or C.ScatterStyleValues.SmoothMarker;

            bool wantsMarker = scatterStyle is C.ScatterStyleValues.Marker
                                            or C.ScatterStyleValues.LineMarker
                                            or C.ScatterStyleValues.SmoothMarker;

            bool smooth      = scatterStyle is C.ScatterStyleValues.Smooth 
                                            or C.ScatterStyleValues.SmoothMarker;

            uint idx = 0;

            foreach (SeriesDefinition seriesDefinition in chartDefinition.Series)
            {
                C.ScatterChartSeries series = new (
                    new C.Index () { Val = idx },
                    new C.Order () { Val = idx },
                    new C.SeriesText(new C.NumericValue() { Text = seriesDefinition.Title })
                );

                // X
                C.NumberLiteral xNum = new ();
                xNum.Append(new C.PointCount() { Val = (uint)seriesDefinition.Points.Count });
                for (int i = 0; i < seriesDefinition.Points.Count; i++)
                    xNum.Append(new C.NumericPoint() { Index = (uint)i, NumericValue = new C.NumericValue(seriesDefinition.Points[i].X?.ToString(CultureInfo.InvariantCulture)) });

                // Y
                C.NumberLiteral yNum = new ();
                yNum.Append(new C.PointCount() { Val = (uint)seriesDefinition.Points.Count });
                for (int i = 0; i < seriesDefinition.Points.Count; i++)
                    yNum.Append(new C.NumericPoint() { Index = (uint)i, NumericValue = new C.NumericValue(seriesDefinition.Points[i].Y?.ToString(CultureInfo.InvariantCulture)) });

                series.Append(new C.Smooth() { Val = smooth });

                if (!wantsLine)
                {
                    ensureShapeLineHidden(series);
                }
                else
                {
                    ensureShapeLineVisible(series);

                    applySeriesFormat(chartDefinition, seriesDefinition, series, toOutline: true);
                }

                if (wantsMarker)
                {
                    ensureMarker(series, chartDefinition, seriesDefinition);
                    applyPointLevelStyling(seriesDefinition, series, useMarker: true);
                }
                else
                {
                    series.Append(new C.Marker(new C.Symbol() { Val = C.MarkerStyleValues.None }));
                }

                series.Append(new C.XValues(xNum));
                series.Append(new C.YValues(yNum));

                addDataLabels(series, seriesDefinition);

                scatterChart.Append(series);

                idx++;
            }
        }
        private     static  void                        addRadarSeries          (ChartDefinition chartDefinition, C.RadarChart radarChart)                          
        {
            uint idx = 0;

            foreach (SeriesDefinition seriesDefinition in chartDefinition.Series)
            {
                C.RadarChartSeries series = new (
                    new C.Index { Val = idx },
                    new C.Order { Val = idx },
                    new C.SeriesText(new C.NumericValue { Text = seriesDefinition.Title })
                );

                buildCategoryAndValues(seriesDefinition, out var cat, out var vals);

                switch (chartDefinition.RadarStyle)
                {
                    case Internal.RadarStyle.Marker  :
                        // çizgi + marker
                        ensureShapeLineVisible(series);
                        ensureMarker     (series, chartDefinition, seriesDefinition);
                        applyPointLevelStyling (seriesDefinition, series, useMarker: true);
                        applySeriesFormat(chartDefinition, seriesDefinition, series, toOutline: true); // outline
                        break;
                    case Internal.RadarStyle.Filled  :
                        // alan dolu, marker istemiyoruz, çizgi opsiyonel (outline)
                        ensureNoMarker   (series);
                        ensureShapeLineVisible(series);

                        // Seri rengini hem doldurma hem outline için uygula
                        // (toOutline=false -> fill; true -> line)
                        applySeriesFormat(chartDefinition, seriesDefinition, series); // fill
                        applySeriesFormat(chartDefinition, seriesDefinition, series, toOutline: true ); // outline
                        break;
                    case Internal.RadarStyle.Standard:
                    default                          :
                        // sadece çizgi, marker yok
                        ensureNoMarker   (series);
                        ensureShapeLineVisible(series);

                        // seri çizgi rengi
                        applySeriesFormat(chartDefinition, seriesDefinition, series, toOutline: true);
                        break;
                }

                series.Append(cat);
                series.Append(vals);

                addDataLabels(series, seriesDefinition);

                radarChart.Append(series);

                idx++;
            }
        }

        private     static  void                        addTitle                (ITitleOwner titleOwner, OpenXmlCompositeElement owner)                             
        {
            if (string.IsNullOrWhiteSpace(titleOwner?.Title) || owner is null)
                return;

            TitleDefinition titleDefinition = titleOwner.Title;

            A.Text      text      = new(titleDefinition.Text);
            A.Run       run       = new(text);
            A.Paragraph paragraph = new(run);
            C.RichText  richText  = new(new A.BodyProperties(), new A.ListStyle(), paragraph);
            C.ChartText charText  = new(richText);
            C.Title     title     = new(charText);
            Overlay     overlay   = new Overlay() { Val = false };

            title.Append(overlay);
            owner.Append(title);

            if (titleDefinition.TextFormat is not null)
                applyTextFormat(richText, titleDefinition.TextFormat);
        }
        private     static  void                        addLayout               (ChartLayoutDefinition layoutDefinition, PlotArea plotArea)                         
        {
            if (plotArea is null)
                return;

            Layout layout;

            if (layoutDefinition is null || !layoutDefinition.HasManualLayout)
            {
                layout = new Layout();
            }
            else
            {
                ManualLayout manualLayout = new();

                if (layoutDefinition.X.HasValue)
                    manualLayout.Append(new C.Left() { Val = layoutDefinition.X.Value / 100.0 });

                if (layoutDefinition.Y.HasValue)
                    manualLayout.Append(new C.Top() { Val = layoutDefinition.Y.Value / 100.0 });

                if (layoutDefinition.Width.HasValue)
                    manualLayout.Append(new C.Width() { Val = layoutDefinition.Width.Value / 100.0 });

                if (layoutDefinition.Height.HasValue)
                    manualLayout.Append(new C.Height() { Val = layoutDefinition.Height.Value / 100.0 });

                if (layoutDefinition.XMode.HasValue)
                    manualLayout.Append(new C.LeftMode()
                    {
                        Val = layoutDefinition.XMode.Value == LayoutMode.Edge
                            ? LayoutModeValues.Edge
                            : LayoutModeValues.Factor
                    });

                if (layoutDefinition.YMode.HasValue)
                    manualLayout.Append(new C.TopMode()
                    {
                        Val = layoutDefinition.YMode.Value == LayoutMode.Edge
                            ? LayoutModeValues.Edge
                            : LayoutModeValues.Factor
                    });

                if (layoutDefinition.Target.HasValue)
                    manualLayout.Append(new C.LayoutTarget()
                    {
                        Val = layoutDefinition.Target.Value == Internal.LayoutTarget.Inner
                            ? LayoutTargetValues.Inner
                            : LayoutTargetValues.Outer
                    });

                layout = new (manualLayout);
            }

            plotArea.InsertAt(layout, 0);
            //plotArea.Append(layout);
        }
        private     static  void                        addCategoryAxis         (ChartDefinition chartDefinition, PlotArea plotArea, uint categoryAxisId, uint valueAxisId, AxisPositionValues position = AxisPositionValues.Bottom) 
        {
            if (chartDefinition is null || plotArea is null)
                return;

            CategoryAxis catAx = new(
                new AxisId              () { Val = categoryAxisId },
                new Delete              () { Val = !chartDefinition.ShowCategoryAxis },
                new Scaling             (new Orientation() { Val = OrientationValues.MinMax }),
                new AxisPosition        () { Val = position },
                new TickLabelPosition   () { Val = TickLabelPositionValues.NextTo },
                new CrossingAxis        () { Val = valueAxisId },
                new Crosses             () { Val = CrossesValues.AutoZero },
                new AutoLabeled         () { Val = true },
                new LabelAlignment      () { Val = LabelAlignmentValues.Center },
                new LabelOffset         () { Val = 100 },
                new C.TickLabelSkip     () { Val = 1 },
                new C.TickMarkSkip      () { Val = 1 },
                new C.NoMultiLevelLabels() { Val = true }
            );

            switch (chartDefinition.Type)
            {
                case ChartType.Bar     :
                case ChartType.Column  :
                case ChartType.Line    :
                case ChartType.Area    :
                case ChartType.Radar   :
                    catAx.InsertAt(new C.MajorTickMark() { Val = C.TickMarkValues.None    }, 3);
                    catAx.InsertAt(new C.MinorTickMark() { Val = C.TickMarkValues.Outside }, 4);
                    break;
                case ChartType.Scatter :
                case ChartType.Bubble  :
                    catAx.InsertAt(new C.MajorTickMark() { Val = C.TickMarkValues.Outside }, 3);
                    catAx.InsertAt(new C.MinorTickMark() { Val = C.TickMarkValues.None    }, 4);
                    break;
                case ChartType.Pie     :
                case ChartType.Doughnut:
                default                :
                    catAx.InsertAt(new C.MajorTickMark() { Val = C.TickMarkValues.None }, 3);
                    catAx.InsertAt(new C.MinorTickMark() { Val = C.TickMarkValues.None }, 4);
                    break;
            }

            applyAxisDefinition(catAx, chartDefinition.CategoryAxis);

            bool addMajorGridLines;
            FormatDefinition targetMajorGridlinesFormat = null;

            switch (chartDefinition.Type)
            {
                case ChartType.Bar     :
                    addMajorGridLines = chartDefinition.Grid?.ShowHorizontal == true;
                    targetMajorGridlinesFormat = chartDefinition.Grid?.HorizontalFormat;
                    break;
                case ChartType.Column  :
                case ChartType.Line    :
                case ChartType.Area    :
                case ChartType.Radar   :
                    addMajorGridLines = chartDefinition.Grid?.ShowVertical == true;
                    targetMajorGridlinesFormat = chartDefinition.Grid?.VerticalFormat;
                    break;
                case ChartType.Pie     :
                case ChartType.Doughnut:
                case ChartType.Bubble  :
                case ChartType.Scatter :
                default                :
                    addMajorGridLines = false;
                    break;
            }

            if (addMajorGridLines)
                addMajorGridlines(catAx, targetMajorGridlinesFormat);

            plotArea.Append(catAx);
        }
        private     static  void                        addValueAxis            (ChartDefinition chartDefinition, PlotArea plotArea, uint categoryAxisId, uint valueAxisId, AxisPositionValues position = AxisPositionValues.Left)   
        {
            if (chartDefinition is null || plotArea is null)
                return;

            ValueAxis    valAx = new (
                new AxisId           () { Val = valueAxisId },
                new Scaling          (new Orientation() { Val = OrientationValues.MinMax }),
                new AxisPosition     () { Val = position },
                new C.MajorTickMark  () { Val = chartDefinition.Type == ChartType.Radar ? C.TickMarkValues.Cross : C.TickMarkValues.Outside },
                new C.MinorTickMark  () { Val = C.TickMarkValues.None },
                new C.NumberingFormat() { FormatCode = "General", SourceLinked = true },
                new TickLabelPosition() { Val = TickLabelPositionValues.NextTo },
                new CrossingAxis     () { Val = categoryAxisId },
                new Crosses          () { Val = CrossesValues.AutoZero },
                new CrossBetween     () { Val = CrossBetweenValues.Between }
            );

            AxisDefinition targetAxisDefinition = chartDefinition.ValueAxis;
            bool           showAxis             = chartDefinition.ShowValueAxis;
            bool           addMajorGridLines;
            FormatDefinition targetMajorGridlinesFormat = null;

            switch (chartDefinition.Type)
            {
                case ChartType.Bar     :
                    addMajorGridLines = chartDefinition.Grid?.ShowVertical == true;
                    targetMajorGridlinesFormat = chartDefinition.Grid?.VerticalFormat;
                    break;
                case ChartType.Column  :
                case ChartType.Line    :
                case ChartType.Area    :
                case ChartType.Radar   :
                    addMajorGridLines = chartDefinition.Grid?.ShowHorizontal == true;
                    targetMajorGridlinesFormat = chartDefinition.Grid?.HorizontalFormat;
                    break;
                case ChartType.Bubble  :
                case ChartType.Scatter :
                    if (position is AxisPositionValues.Left or AxisPositionValues.Right)
                    {
                        addMajorGridLines = chartDefinition.Grid?.ShowHorizontal == true;
                        targetMajorGridlinesFormat = chartDefinition.Grid?.HorizontalFormat;
                    }
                    else
                    {
                        addMajorGridLines    = chartDefinition.Grid?.ShowVertical == true;
                        targetAxisDefinition = chartDefinition.CategoryAxis;
                        showAxis             = chartDefinition.ShowCategoryAxis;
                        targetMajorGridlinesFormat = chartDefinition.Grid?.VerticalFormat;
                    }
                    break;
                case ChartType.Pie     :
                case ChartType.Doughnut:
                default                :
                    addMajorGridLines = false;
                    break;
            }

            applyAxisDefinition(valAx, targetAxisDefinition);

            if (addMajorGridLines)
                addMajorGridlines(valAx, targetMajorGridlinesFormat);

            valAx.InsertAt(new Delete() { Val = !showAxis }, 2);

            plotArea.Append(valAx);
        }
        private     static  void                        addDataLabels           (OpenXmlCompositeElement owner, IValueLabelsContainer valueLabelsContainer)         
        {
            if (owner is null || valueLabelsContainer is null)
                return;

            ChartDefinition chartDefinition = valueLabelsContainer.GetChartDefinition();

            if (chartDefinition is null)
                return;

            ValueLabelDefinition valueLabelsDefinition = valueLabelsContainer.ValueLabels;

            if (valueLabelsDefinition is null || !valueLabelsDefinition.Show)
                return;

            C.DataLabels dataLabels = new ();

            dataLabels.Append(new C.ShowValue() { Val = true });

            DataLabelPositionValues? position = mapDataLabelPosition(valueLabelsDefinition.Position, chartDefinition);

            if (position is not null)
            {
                C.DataLabelPosition dataLabelPosition = new C.DataLabelPosition() { Val = position.Value };
                dataLabels.Append(dataLabelPosition);
            }

            bool showPercent     = false;
            bool showBubbleSize  = false;
            bool showLeaderLines = false;

            switch (chartDefinition.Type)
            {
                case ChartType.Pie     :
                    showPercent     = valueLabelsDefinition.ShowPercent;
                    showLeaderLines = position == DataLabelPositionValues.OutsideEnd || position == DataLabelPositionValues.BestFit;
                    break;
                case ChartType.Doughnut:
                    showPercent     = valueLabelsDefinition.ShowPercent;
                    showLeaderLines = true;
                    break;
                case ChartType.Scatter :
                    break;
                case ChartType.Bubble  :
                    showBubbleSize = valueLabelsDefinition.ShowBubbleSize;
                    break;
                case ChartType.Radar   :
                    break;
                default                :
                    break;
            }

            dataLabels.Append(new C.ShowLegendKey    () { Val = valueLabelsDefinition.ShowLegendKey   });
            dataLabels.Append(new C.ShowCategoryName () { Val = valueLabelsDefinition.ShowCategoryName});
            dataLabels.Append(new C.ShowSeriesName   () { Val = valueLabelsDefinition.ShowSeriesName  });
            dataLabels.Append(new C.ShowPercent      () { Val = showPercent                           });
            dataLabels.Append(new C.ShowBubbleSize   () { Val = showBubbleSize                        });
            dataLabels.Append(new C.ShowLeaderLines  () { Val = showLeaderLines                       });

            if (valueLabelsDefinition.TextFormat is not null)
            {
                C.TextProperties textProperties = new();

                if (applyTextFormat(textProperties, valueLabelsDefinition.TextFormat))
                    dataLabels.Append(textProperties);
            }

            if (valueLabelsContainer is SeriesDefinition seriesDefinition && seriesDefinition.Points.Any(x => x.ValueLabels is not null))
            {
                uint pointIndex = 0;

                foreach (PointDefinition pointDefinition in seriesDefinition.Points)
                {
                    if (pointDefinition.ValueLabels is ValueLabelDefinition pointValueLabelDefinition)
                    {
                        C.DataLabel pointDataLabel = new (
                            new C.Index           () { Val = pointIndex                                 },
                            new C.ShowValue       () { Val = true                                       },
                            new C.ShowLegendKey   () { Val = pointValueLabelDefinition.ShowLegendKey    },
                            new C.ShowCategoryName() { Val = pointValueLabelDefinition.ShowCategoryName },
                            new C.ShowSeriesName  () { Val = pointValueLabelDefinition.ShowSeriesName   },
                            new C.ShowPercent     () { Val = pointValueLabelDefinition.ShowPercent && (chartDefinition.Type == ChartType.Pie || chartDefinition.Type == ChartType.Doughnut) },
                            new C.ShowBubbleSize  () { Val = pointValueLabelDefinition.ShowBubbleSize && chartDefinition.Type == ChartType.Bubble }
                        );

                        if (pointValueLabelDefinition.Position is not null)
                        {
                            DataLabelPositionValues? pointLabelPosition = mapDataLabelPosition(pointValueLabelDefinition.Position.Value, chartDefinition);

                            if (pointLabelPosition is not null)
                            {
                                C.DataLabelPosition dataLabelPosition = new () { Val = pointLabelPosition.Value };
                                pointDataLabel.Append(dataLabelPosition);
                            }
                        }

                        if (pointValueLabelDefinition.TextFormat is not null)
                        {
                            C.TextProperties textProperties = new ();

                            if (applyTextFormat(textProperties, pointValueLabelDefinition.TextFormat))
                                pointDataLabel.Append(textProperties);
                        }

                        dataLabels.Append(pointDataLabel);
                    }

                    pointIndex++;
                }

            }

            owner.Append(dataLabels);
        }
        private     static  void                        addMajorGridlines       (OpenXmlCompositeElement owner, FormatDefinition gridFormatDefinition)              
        {
            if (owner is null)
                return;

            MajorGridlines majorGridlines = new();

            if (gridFormatDefinition is not null)
                applyFormat(majorGridlines, gridFormatDefinition);

            owner.Append(majorGridlines);
        }
        private     static  void                        addAxisIds              (OpenXmlCompositeElement owner, params uint[] ids)                                  
        {
            if (owner is null || ids is null) 
                return;

            foreach (uint id in ids)
                owner.Append(new C.AxisId { Val = id });
        }
        private     static  void                        add3DView               (C.Chart chart, ChartDefinition chartDefinition)                                
        {

            if (chart is null || chartDefinition?.ThreeDView is not ThreeDViewDefinition viewDefinition) 
                return;


            bool addView3D = viewDefinition.RotationX.HasValue                               
                          || viewDefinition.RotationY.HasValue                               
                          || viewDefinition.DepthPercent.HasValue                            
                          || viewDefinition.HeightPercent.HasValue                           
                          || viewDefinition.RightAngleAxes.HasValue                          
                          || (viewDefinition.Perspective.HasValue && viewDefinition.RightAngleAxes != true);

            if (addView3D)
            {
                C.View3D view3D = new ();

                // rotX: SByteValue, -90..90
                if (viewDefinition.RotationX.HasValue)
                    view3D.Append(new C.RotateX { Val = new SByteValue((sbyte)clamp(viewDefinition.RotationX.Value, -90, 90)) });

                // rotY: UInt16Value, 0..360
                if (viewDefinition.RotationY.HasValue)
                    view3D.Append(new C.RotateY { Val = new UInt16Value((ushort)clamp(viewDefinition.RotationY.Value, 0, 360)) });

                // depthPercent: UInt16Value, 20..2000
                if (viewDefinition.DepthPercent.HasValue)
                    view3D.Append(new C.DepthPercent { Val = new UInt16Value((ushort)clamp(viewDefinition.DepthPercent.Value, 20, 2000)) });

                // hPercent: UInt16Value, 5..500
                if (viewDefinition.HeightPercent.HasValue)
                    view3D.Append(new C.HeightPercent { Val = new UInt16Value((ushort)clamp(viewDefinition.HeightPercent.Value, 5, 500)) });

                // rAngAx: bool
                if (viewDefinition.RightAngleAxes.HasValue)
                    view3D.Append(new C.RightAngleAxes { Val = viewDefinition.RightAngleAxes.Value });

                // perspective: ByteValue, 0..240
                if (viewDefinition.Perspective.HasValue && viewDefinition.RightAngleAxes != true)
                    view3D.Append(new C.Perspective { Val = new ByteValue((byte)clamp(viewDefinition.Perspective.Value, 0, 240)) });

                // Sıra: view3D, plotArea’dan önce olmalı
                PlotArea plotArea = chart.Elements<PlotArea>().FirstOrDefault();

                if (plotArea is not null)
                    chart.InsertBefore(view3D, plotArea);
                else
                    chart.Append(view3D);
            }

            switch (chartDefinition.Type)
            {
                case ChartType.Bar   :
                case ChartType.Column:
                case ChartType.Line  :
                case ChartType.Area  :
                    add3DFloorAndWalls(chart, viewDefinition);
                    break;
                default:
                    break;
            }
        }
        private     static  void                        add3DFloorAndWalls      (C.Chart chart, ThreeDViewDefinition viewDefinition)                                
        {
            if (chart is null || viewDefinition is null) 
                return;

            static  T getWallOrFloor<T> (bool show, FormatDefinition formatDefinition, FormatDefinition defaultFormatDefinition) where T : OpenXmlCompositeElement, new()
            {
                T elem = new T();

                if (!show)
                {
                    // görünmez yapmak için: NoFill + NoLine
                    elem.Append(new C.ChartShapeProperties(
                        new A.NoFill(),
                        new A.Outline(new A.NoFill())
                    ));
                }
                else
                {
                    applyFormat(elem, formatDefinition, defaultFormatDefinition);
                }

                return elem;
            }

            if (viewDefinition.ShowFloor.HasValue)
                chart.Append(getWallOrFloor<C.Floor>(viewDefinition.ShowFloor.Value, viewDefinition.FloorFormat, viewDefinition.DefaultFormat));

            if (viewDefinition.ShowBackWall.HasValue)
                chart.Append(getWallOrFloor<C.BackWall>(viewDefinition.ShowBackWall.Value, viewDefinition.BackWallFormat, viewDefinition.DefaultFormat));

            if (viewDefinition.ShowSideWall.HasValue)
                chart.Append(getWallOrFloor<C.SideWall>(viewDefinition.ShowSideWall.Value, viewDefinition.SideWallFormat, viewDefinition.DefaultFormat));
        }

        private     static  void                        applyFormat             (OpenXmlCompositeElement node, params FormatDefinition[] formats)                   
        {
            if (node is null || formats is null)
                return;

            C.ChartShapeProperties charShapeProperties = getOrAddChartShapeProps(node);

            if (charShapeProperties is null) 
                return;

            FillDefinition    getFillDefinition  ()
            {
                for (int i = 0; i < formats.Length; i++) 
                    if (formats[i]?.FillDefinition is FillDefinition fd)
                        return fd;

                return null;
            }
            LineDefinition    getLineDefinition  ()
            {
                for (int i = 0; i < formats.Length; i++)
                    if (formats[i]?.LineDefinition is LineDefinition ld)
                        return ld;

                return null;
            }
            EffectsDefinition getEffectsDefintion()
            {
                for (int i = 0; i < formats.Length; i++)
                    if (formats[i]?.EffectsDefinition is EffectsDefinition ed)
                        return ed;

                return null;
            }

            if (getFillDefinition() is FillDefinition fillDefinition)
            {
                charShapeProperties.RemoveAllChildren<A.NoFill>();
                charShapeProperties.RemoveAllChildren<A.SolidFill>();
                charShapeProperties.RemoveAllChildren<A.PatternFill>();
                charShapeProperties.RemoveAllChildren<A.GradientFill>();

                bool fillApplied = false;

                if (fillDefinition.GradientFillDefinition is not null)
                {
                    A.GradientFill gradientFill = buildGradientFill(fillDefinition.GradientFillDefinition);

                    if (gradientFill is not null)
                    {
                        charShapeProperties.Append(gradientFill);
                        fillApplied = true;
                    }
                }
                
                if (!fillApplied && fillDefinition.PatternFillDefinition is not null)
                {
                    A.PatternFill patternFill = buildPatternFill(fillDefinition.PatternFillDefinition);

                    if (patternFill is not null)
                    {
                        charShapeProperties.Append(patternFill);
                        fillApplied = true;
                    }
                }

                if (!fillApplied && fillDefinition.SolidFillDefinition is not null)
                {
                    A.SolidFill solidFill = buildSolidFill(fillDefinition.SolidFillDefinition.Color, "dedede");

                    if (solidFill is not null)
                    {
                        charShapeProperties.Append(solidFill);
                        fillApplied = true;
                    }
                }
            }

            // LINE (A.Outline)
            if (getLineDefinition() is LineDefinition lineDefinition)
            {
                A.Outline outline = getOrAddOutline(charShapeProperties);

                // visibility
                if (!lineDefinition.Visible)
                {
                    outline.RemoveAllChildren();
                    outline.Append(new A.NoFill());
                }
                else
                {
                    if (lineDefinition.Width is not null)
                        outline.Width = lineDefinition.Width.ToEmu();

                    // stroke (solid/gradient)
                    outline.RemoveAllChildren<A.SolidFill>();
                    outline.RemoveAllChildren<A.GradientFill>();
                    outline.RemoveAllChildren<A.NoFill>();

                    bool styleApplied = false;

                    if (lineDefinition.GradientLineDefinition is not null)
                    {
                        A.GradientFill gradientFill = buildGradientFill(lineDefinition.GradientLineDefinition);

                        if (gradientFill is not null)
                        {
                            outline.Append(gradientFill);
                            styleApplied = true;
                        }
                    }

                    if (!styleApplied && lineDefinition.SolidLineDefinition is not null)
                    {
                        A.SolidFill solidFill = buildSolidFill(lineDefinition.SolidLineDefinition.Color, "000000");

                        if (solidFill is not null)
                        {
                            outline.Append(solidFill);
                            styleApplied = true;
                        }
                    }

                    // stroke knobs
                    if (lineDefinition.DashPreset is not null)       
                        outline.Append(new A.PresetDash() { Val = mapDash(lineDefinition.DashPreset.Value) });

                    if (lineDefinition.CompoundPreset is not null)
                        outline.CompoundLineType = mapCompound(lineDefinition.CompoundPreset.Value);

                    if (lineDefinition.CapPreset is not null)
                        outline.CapType = mapLineCap(lineDefinition.CapPreset.Value);

                    // join
                    applyFormatJoin(outline, lineDefinition);

                    // ends
                    if (lineDefinition.Begin_Arrow_Type is not null)
                        outline.Append(new A.HeadEnd{
                            Type   = mapLineEndType(lineDefinition.Begin_Arrow_Type.Value),
                            Width  = mapLineEndWidth(lineDefinition.Begin_Arrow_Width ?? LineEndWidthPreset.Medium),
                            Length = mapLineEndLength(lineDefinition.Begin_Arrow_Length ?? LineEndLengthPreset.Medium)
                        });

                    if (lineDefinition.End_Arrow_Type is not null)
                        outline.Append(new A.TailEnd{
                            Type   = mapLineEndType(lineDefinition.End_Arrow_Type.Value),
                            Width  = mapLineEndWidth(lineDefinition.End_Arrow_Width ?? LineEndWidthPreset.Medium),
                            Length = mapLineEndLength(lineDefinition.End_Arrow_Length ?? LineEndLengthPreset.Medium)
                        });
                }
            }

            // EFFECTS (shadow, glow, soft-edges, 3D, reflection)
            if (getEffectsDefintion() is EffectsDefinition effectsDefinition)
                applyEffects(node, effectsDefinition);
        }
        private     static  void                        applyEffects            (OpenXmlCompositeElement node, EffectsDefinition effectsDefinition)                 
        {
            if (node is null || effectsDefinition is null)
                return;

            C.ChartShapeProperties chartShapeProperties = getOrAddChartShapeProps(node);

            if (chartShapeProperties is null)
                return;

            A.EffectList effectList           = new();
            Dimension    placeHolderDimension = new (4, UnitType.Pt);

            // Önce Glow, Sonra Shadow, aksi halde file corrupt !!!

            if (effectsDefinition.GlowDefinition is GlowDefinition glowDefinition)
                applyGlow(effectList, glowDefinition, placeHolderDimension);

            if (effectsDefinition.ShadowDefinition is ShadowDefinition shadowDefinition)
                applyShadow(effectList, shadowDefinition, placeHolderDimension, ShadowType.Inner);

            if (effectsDefinition.SoftEdgesDefinition is SoftEdgesDefinition softEdgesDefinition)
                applySoftEdges(effectList, softEdgesDefinition, placeHolderDimension);

            if (effectsDefinition.ReflectionDefinition is ReflectionDefinition reflectionDefinition)
                applyReflection(effectList, reflectionDefinition);

            if (effectsDefinition.Format3dDefinition is not null)
            {
                applyFormat3D(chartShapeProperties, effectsDefinition.Format3dDefinition);
            }

            if (effectList.HasChildren)
            {
                chartShapeProperties.RemoveAllChildren<A.EffectList>();
                chartShapeProperties.Append(effectList);
            }
        }
        private     static  void                        applyFormat3D           (C.ChartShapeProperties shapeProperties, Format3DDefinition format3d)               
        {
            if (shapeProperties is null || format3d is null) 
                return;

            LightingDefinition light = format3d.LightingDefinition;

            //Önce Scene3DType, Sonra Shape3DType !!!

            // a:scene3d (camera + lightRig)
            shapeProperties.RemoveAllChildren<A.Scene3DType>();
            A.Scene3DType scene = shapeProperties.AppendChild(new A.Scene3DType());

            // Camera (basit preset)
            scene.Camera = new A.Camera { Preset = A.PresetCameraValues.OrthographicFront };

            // LightRig
            A.LightRig lr = new ()
            {
                Rig = light?.Preset is null ? A.LightRigValues.ThreePoints : mapLightPreset(light.Preset.Value),
                Direction = light?.Direction is null ? A.LightRigDirectionValues.Top : mapLightDir(light.Direction.Value)
            };

            if (light?.Angle is not null)
                lr.Append(new A.Rotation { Latitude = 0, Longitude = 0, Revolution = light.Angle.Value * 60000 });

            scene.LightRig = lr;

            // a:sp3d (malzeme + bevel)
            shapeProperties.RemoveAllChildren<A.Shape3DType>();
            A.Shape3DType sp3d = shapeProperties.AppendChild(new A.Shape3DType());

            if (format3d.Material is not null)
                sp3d.PresetMaterial = mapMaterial(format3d.Material.Value); // A.PresetMaterialTypeValues

            if (format3d.BevelDefinition is BevelDefinition bevelDefinition && bevelDefinition.HasData)
            {
                if (bevelDefinition.TopWidth is not null || bevelDefinition.TopHeight is not null || bevelDefinition.TopPreset is not null)
                    sp3d.BevelTop = new A.BevelTop
                    {
                        Width  = bevelDefinition.TopWidth?.ToEmu(),
                        Height = bevelDefinition.TopHeight?.ToEmu(),
                        Preset = bevelDefinition.TopPreset is null ? null : mapBevel(bevelDefinition.TopPreset.Value)
                    };
                if (bevelDefinition.BottomWidth is not null || bevelDefinition.BottomHeight is not null || bevelDefinition.BottomPreset is not null)
                    sp3d.BevelBottom = new A.BevelBottom
                    {
                        Width  = bevelDefinition.BottomWidth?.ToEmu(),
                        Height = bevelDefinition.BottomHeight?.ToEmu(),
                        Preset = bevelDefinition.BottomPreset is null ? null : mapBevel(bevelDefinition.BottomPreset.Value)
                    };
            }
        }
        private     static  void                        applyFormatJoin         (A.Outline outline, LineDefinition lineDefinition)                                  
        {
            if (outline is null)
                return;

            outline.RemoveAllChildren<A.Miter>();
            outline.RemoveAllChildren<A.Round>();
            outline.RemoveAllChildren<A.Bevel>();

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
                    outline.Append(new A.Miter(){ Limit = (lineDefinition.JoinMiterLimit ?? 800000) });
                    break;
                default:
                    break;
            }
        }
        private     static  bool                        applyTextFormat         (OpenXmlCompositeElement txPrOwner, TextFormatDefinition textFormatDefinition)      
        {
            if (txPrOwner is null || textFormatDefinition is null) 
                return false;

            OpenXmlCompositeElement txPr = txPrOwner;

            A.BodyProperties bodyPr = txPr.GetFirstChild<A.BodyProperties>() ?? txPr.AppendChild(new A.BodyProperties());

            if (txPr.GetFirstChild<A.ListStyle>() is null)
                txPr.AppendChild(new A.ListStyle());

            A.Paragraph p = txPr.GetFirstChild<A.Paragraph>() ?? txPr.AppendChild(new A.Paragraph());

            A.ParagraphProperties pPr = p.GetFirstChild<A.ParagraphProperties>() ?? p.PrependChild(new A.ParagraphProperties());

            A.DefaultRunProperties defRPr = pPr.GetFirstChild<A.DefaultRunProperties>() ?? pPr.AppendChild(new A.DefaultRunProperties());

            if (textFormatDefinition.AlignHorizontal is not null)
                pPr.Alignment = mapTextAlign(textFormatDefinition.AlignHorizontal.Value);

            if (textFormatDefinition.AlignVertical is not null)
                bodyPr.Anchor = mapTextAnchor(textFormatDefinition.AlignVertical.Value);

            if (textFormatDefinition.Wrap is not null)
                bodyPr.Wrap = mapTextWrap(textFormatDefinition.Wrap.Value);

            if (textFormatDefinition.MarginLeft   is not null) 
                bodyPr.LeftInset   = textFormatDefinition.MarginLeft.ToEmu();

            if (textFormatDefinition.MarginRight  is not null) 
                bodyPr.RightInset  = textFormatDefinition.MarginRight.ToEmu();

            if (textFormatDefinition.MarginTop    is not null) 
                bodyPr.TopInset    = textFormatDefinition.MarginTop.ToEmu();

            if (textFormatDefinition.MarginBottom is not null) 
                bodyPr.BottomInset = textFormatDefinition.MarginBottom.ToEmu();

            if (textFormatDefinition.Rotate is not null)
                bodyPr.Rotation = textFormatDefinition.Rotate.Value * 60000;

            applyTextOutline(defRPr, textFormatDefinition.TextOutlineDefinition);
            applyTextFill   (defRPr, textFormatDefinition.TextFillDefinition);
            applyTextEffects(defRPr, textFormatDefinition.EffectsDefinition);
            applyFont       (defRPr, textFormatDefinition.FontDefinition);

            if (p.GetFirstChild<A.Run>() is null)
                p.Append(new A.Run(new A.RunProperties(), new A.Text() { Text = string.Empty }));

            return true;
        }
        private     static  void                        applyFont               (A.TextCharacterPropertiesType runProperties, FontDefinition fontDefinition)        
        {
            if (runProperties is null || fontDefinition is null) 
                return;

            if (!string.IsNullOrWhiteSpace(fontDefinition.Family))
            {
                runProperties.RemoveAllChildren<A.LatinFont>();
                runProperties.Append(new A.LatinFont { Typeface = fontDefinition.Family });
            }

            if (fontDefinition.Size is not null)
                runProperties.FontSize = (int)Math.Round(fontDefinition.Size.Value * 100.0);

            if (fontDefinition.Bold is not null) 
                runProperties.Bold = fontDefinition.Bold;

            if (fontDefinition.Italic is not null) 
                runProperties.Italic = fontDefinition.Italic;

            if (fontDefinition.Underline is not null)
                runProperties.Underline = mapUnderline(fontDefinition.Underline.Value);

            if (fontDefinition.Strike is not null)
                runProperties.Strike = mapStrike(fontDefinition.Strike.Value);

            if (fontDefinition.Caps is not null)
                runProperties.Capital = mapCaps(fontDefinition.Caps.Value);

            if (fontDefinition.Kerning is not null)
                runProperties.Kerning = (int)Math.Max(0, Math.Round(fontDefinition.Kerning.Value * 100.0)); // 1/100 pt

            if (fontDefinition.Spacing is not null)
                runProperties.Spacing = (int)Math.Round(fontDefinition.Spacing.Value); // DrawingML spacing: ±EMUs değil; text char spacing (1/1000 em). İstersen burada 1/1000 em dönüşümü kurala bağlayabilirsin.
        }
        private     static  void                        applyTextFill           (A.TextCharacterPropertiesType runProperties, FillDefinition fillDefinition)        
        {
            if (runProperties is null) 
                return;

            runProperties.RemoveAllChildren<A.NoFill      >();
            runProperties.RemoveAllChildren<A.SolidFill   >();
            runProperties.RemoveAllChildren<A.GradientFill>();
            runProperties.RemoveAllChildren<A.PatternFill >();
            runProperties.RemoveAllChildren<A.BlipFill    >();
            runProperties.RemoveAllChildren<A.GroupFill   >();

            if (fillDefinition is null) 
                return;

            if (fillDefinition.GradientFillDefinition is not null)
            {
                A.GradientFill gf = buildGradientFill(fillDefinition.GradientFillDefinition);

                if (gf is not null) 
                { 
                    runProperties.Append(gf); 
                    return; 
                }
            }

            if (fillDefinition.PatternFillDefinition is not null)
            {
                A.PatternFill pf = buildPatternFill(fillDefinition.PatternFillDefinition);

                if (pf is not null) 
                { 
                    runProperties.Append(pf); 
                    return; 
                }
            }

            if ( fillDefinition.SolidFillDefinition?.Color is not null)
            {
                A.SolidFill sf = new (toHexFill(fillDefinition.SolidFillDefinition.Color, "000000"));

                if (sf is not null)
                    runProperties.Append(sf);
            }
        }
        private     static  void                        applyTextOutline        (A.TextCharacterPropertiesType runProperties, LineDefinition lineDefinition)        
        {
            if (runProperties is null || lineDefinition is null) 
                return;

            A.Outline outline = getOrAddOutline(runProperties);

            outline.RemoveAllChildren<A.NoFill      >();
            outline.RemoveAllChildren<A.SolidFill   >();
            outline.RemoveAllChildren<A.GradientFill>();
            outline.RemoveAllChildren<A.PatternFill >();
            outline.RemoveAllChildren<A.BlipFill    >();
            outline.RemoveAllChildren<A.GroupFill   >();

            if (lineDefinition.Width is not null) 
                outline.Width = lineDefinition.Width.ToEmu();

            bool fillApplied = false;

            if (lineDefinition.GradientLineDefinition is not null)
            {
                A.GradientFill gradientFill = buildGradientFill(lineDefinition.GradientLineDefinition);

                if (gradientFill is not null)
                {
                    outline.Append(gradientFill);
                    fillApplied = true;
                }
            }
            
            if (!fillApplied && lineDefinition.SolidLineDefinition is not null)
            {
                A.SolidFill solidFill = buildSolidFill(lineDefinition.SolidLineDefinition.Color, "000000");
                
                if (solidFill is not null)
                {
                    outline.Append(solidFill);
                    fillApplied = true;
                }
            }

            if (!fillApplied && !lineDefinition.Visible)
                outline.Append(new A.NoFill());

            if (lineDefinition.DashPreset is not null) 
            { 
                outline.RemoveAllChildren<A.PresetDash>(); 
                outline.Append(new A.PresetDash { Val = mapDash(lineDefinition.DashPreset.Value) }); 
            }

            if (lineDefinition.CompoundPreset is not null) 
                outline.CompoundLineType = mapCompound(lineDefinition.CompoundPreset.Value);

            if (lineDefinition.CapPreset is not null) 
                outline.CapType = mapLineCap(lineDefinition.CapPreset.Value);

            applyFormatJoin(outline, lineDefinition);

            if (lineDefinition.Begin_Arrow_Type is not null || lineDefinition.Begin_Arrow_Width is not null || lineDefinition.Begin_Arrow_Length is not null)
            {
                outline.RemoveAllChildren<A.HeadEnd>();
                outline.Append(new A.HeadEnd {
                    Type   = lineDefinition.Begin_Arrow_Type   is null ? null : mapLineEndType  (lineDefinition.Begin_Arrow_Type.Value),
                    Width  = lineDefinition.Begin_Arrow_Width  is null ? null : mapLineEndWidth (lineDefinition.Begin_Arrow_Width.Value),
                    Length = lineDefinition.Begin_Arrow_Length is null ? null : mapLineEndLength(lineDefinition.Begin_Arrow_Length.Value),
                });
            }

            if (lineDefinition.End_Arrow_Type is not null || lineDefinition.End_Arrow_Width is not null || lineDefinition.End_Arrow_Length is not null)
            {
                outline.RemoveAllChildren<A.TailEnd>();
                outline.Append(new A.TailEnd {
                    Type   = lineDefinition.End_Arrow_Type   is null ? null : mapLineEndType  (lineDefinition.End_Arrow_Type.Value),
                    Width  = lineDefinition.End_Arrow_Width  is null ? null : mapLineEndWidth (lineDefinition.End_Arrow_Width.Value),
                    Length = lineDefinition.End_Arrow_Length is null ? null : mapLineEndLength(lineDefinition.End_Arrow_Length.Value),
                });
            }
        }
        private     static  void                        applyTextEffects        (A.TextCharacterPropertiesType runProperties, EffectsDefinition effectsDefinition)  
        {
            if (runProperties is null || effectsDefinition is null) 
                return;

            A.EffectList effectList           = new();
            Dimension    placeHolderDimension = new(1, UnitType.Pt);

            // Önce Glow, Sonra Shadow, aksi halde file corrupt !!!

            if (effectsDefinition.GlowDefinition is GlowDefinition glowDefinition)
                applyGlow(effectList, glowDefinition, placeHolderDimension);

            if (effectsDefinition.ShadowDefinition is ShadowDefinition shadowDefinition)
                applyShadow(effectList, shadowDefinition, placeHolderDimension, ShadowType.Outer);

            if (effectsDefinition.SoftEdgesDefinition is SoftEdgesDefinition softEdgesDefinition)
                applySoftEdges(effectList, softEdgesDefinition, placeHolderDimension);

            if (effectsDefinition.ReflectionDefinition is ReflectionDefinition reflectionDefinition)
                applyReflection(effectList, reflectionDefinition);

            if (effectList.HasChildren)
            {
                runProperties.RemoveAllChildren<A.EffectList>();
                runProperties.Append(effectList);
            }
        }
        private     static  void                        applyAxisDefinition     (OpenXmlCompositeElement axisNode, AxisDefinition axisDefinition)                   
        {
            if (axisDefinition is null || axisNode is null)
                return;

            addTitle(axisDefinition, axisNode);

            Scaling scaling = axisNode.Elements<Scaling>().FirstOrDefault();
            if (scaling is null)
            {
                scaling = new Scaling();
                axisNode.PrependChild(scaling); // scaling child'ı başa alıyoruz, genelde ilk sırada olur
            }

            if (axisDefinition.Min.HasValue)
                scaling.Append(new C.MinAxisValue() { Val = axisDefinition.Min.Value });

            if (axisDefinition.Max.HasValue && (!axisDefinition.Min.HasValue || axisDefinition.Min.Value <= axisDefinition.Max.Value))
                scaling.Append(new C.MaxAxisValue() { Val = axisDefinition.Max.Value });

            if (axisDefinition.MajorUnit.HasValue && axisNode is ValueAxis)
                axisNode.Append(new MajorUnit() { Val = axisDefinition.MajorUnit.Value });

            if (axisDefinition.AxisLineFormat is not null)
            {
                applyFormat(axisNode, axisDefinition.AxisLineFormat);
            }

            if (axisDefinition.TickLabelTextFormat is not null)
            {
                C.TextProperties textProperties = getOrAddTextProperties(axisNode);
                applyTextFormat(textProperties, axisDefinition.TickLabelTextFormat);
            }

        }
        private     static  void                        applyBar3DOptions       (C.Bar3DChart bar3DChart, ThreeDViewDefinition viewDefinition)                      
        {
            if (bar3DChart is null || viewDefinition is null) 
                return;

            // gapWidth: 0..500
            bar3DChart.Append(new C.GapWidth() { Val = (UInt16Value)(ushort)clamp(viewDefinition.GapWidth ?? 150, 0, 500) });

            // gapDepth: 0..500
            bar3DChart.Append(new C.GapDepth() { Val = (UInt16Value)(ushort)clamp(viewDefinition.GapDepth ?? 150, 0, 500) });

            // shape: box / cone / coneToMax / cylinder / pyramid / pyramidToMaximum
            bar3DChart.Append(new C.Shape() { Val = mapShapeValues(viewDefinition.Shape) });
        }
        private     static  void                        applyPointLevelStyling  (SeriesDefinition seriesDefinition, OpenXmlCompositeElement series, bool useMarker)                                             
        {
            // useMarker == true  -> dPt/marker/spPr altında uygula (Line/Scatter)
            // useMarker == false -> dPt/spPr altında uygula (Bar/Column/Area/…)

            if (series is null || seriesDefinition?.Points is null) 
                return;

            if (!seriesDefinition.Points.Any(p => p.HasStyle))
                return;

            // dPt’leri cat/val’den ÖNCE yazabilmek için InsertAt kullanacağız,
            // o yüzden burada sadece dPt’yi üretelim:
            for (int i = 0; i < seriesDefinition.Points.Count; i++)
                series.Append(buildDataPoint(i, seriesDefinition.Points[i], useMarker));
        }        
        private     static  void                        applySeriesFormat       (ChartDefinition chartDefinition, SeriesDefinition seriesDefinition, OpenXmlCompositeElement series, bool toOutline = false)    
        {
            if (chartDefinition is null || seriesDefinition is null || series is null || seriesDefinition.Format is null /*|| chartDefinition.VaryColors == true*/)
                return;

            applyFormat(series, seriesDefinition.Format, chartDefinition.SeriesDefaultFormat);
        }
        private     static  void                        applyShadow             (A.EffectList effectList, ShadowDefinition shadowDefinition, Dimension placeHolderDimension, ShadowType defaultShadowType)      
        {
            if (effectList is null || shadowDefinition is null || placeHolderDimension is null)
                return;

            ShadowType shadowType = shadowDefinition.Type ?? defaultShadowType;

            effectList.RemoveAllChildren<A.OuterShadow>();
            effectList.RemoveAllChildren<A.InnerShadow>();
            effectList.RemoveAllChildren<A.PresetShadow>();

            A.RgbColorModelHex   shadowHex   = toHexFill(shadowDefinition?.Color, "000000");
            Int64Value           blurRadius  = (shadowDefinition.BlurRadius ?? placeHolderDimension).ToEmu();
            Int64Value           distance    = (shadowDefinition.Distance   ?? placeHolderDimension).ToEmu();
            Int32Value           direction   = (shadowDefinition.Angle ?? 45) * 60000;
            A.PresetShadowValues preset      = mapShadowPreset(shadowDefinition.Preset);

            if (shadowHex is not null)
            {
                switch (shadowType)
                {
                    case ShadowType.Preset:
                        effectList.Append(new A.PresetShadow(shadowHex) { Distance = distance, Direction = direction, Preset = preset });
                        break;
                    case ShadowType.Outer:
                        effectList.Append(new A.OuterShadow(shadowHex) { BlurRadius = blurRadius, Distance = distance, Direction = direction });
                        break;
                    case ShadowType.Inner:
                    default:
                        effectList.Append(new A.InnerShadow(shadowHex) { BlurRadius = blurRadius, Distance = distance, Direction = direction });
                        break;
                }
            }
        }
        private     static  void                        applyGlow               (A.EffectList effectList, GlowDefinition glowDefinition, Dimension placeHolderDimension)                                        
        {
            if (effectList is null || glowDefinition is null || placeHolderDimension is null)
                return;

            effectList.RemoveAllChildren<A.Glow>();

            A.Glow glow = new()
            {
                RgbColorModelHex = toHexFill(glowDefinition.Color, "000000"),
                Radius           = (glowDefinition.Size ?? placeHolderDimension).ToEmu(),
            };

            effectList.Append(glow);
        }
        private     static  void                        applySoftEdges          (A.EffectList effectList, SoftEdgesDefinition softEdgesDefinition, Dimension placeHolderDimension)                              
        {
            if (effectList is null || softEdgesDefinition is null || placeHolderDimension is null)
                return;

            effectList.RemoveAllChildren<A.SoftEdge>();
            effectList.Append(new A.SoftEdge() { Radius = (softEdgesDefinition.Size ?? placeHolderDimension).ToEmu() });
        }
        private     static  void                        applyReflection         (A.EffectList effectList, ReflectionDefinition reflectionDefinition)                                                            
        {
            if (effectList is null || reflectionDefinition is null || !reflectionDefinition.HasData)
                return;

            effectList.RemoveAllChildren<A.Reflection>();

            A.Reflection reflection = new ();

            if (reflectionDefinition.Blur              is not null) reflection.BlurRadius    = reflectionDefinition.Blur.ToEmu();
            if (reflectionDefinition.Distance          is not null) reflection.Distance      = reflectionDefinition.Distance.ToEmu();
            if (reflectionDefinition.StartTransparency is int st  ) reflection.StartOpacity  = transparencyToAlpha(st);
            if (reflectionDefinition.EndTransparency   is int et  ) reflection.EndAlpha      = transparencyToAlpha(et);
            if (reflectionDefinition.StartPosition     is int sp  ) reflection.StartPosition = sp * 1000;
            if (reflectionDefinition.EndPosition       is int ep  ) reflection.EndPosition   = ep * 1000;

            effectList.Append(reflection);
        }

        private     static  A.SolidFill                 buildSolidFill          (ColorA colorA, string fallbackColor=null)                                          
        {
            if (colorA is null && string.IsNullOrWhiteSpace(fallbackColor))
                return null;

            A.RgbColorModelHex rgb = toHexFill(colorA, fallback: fallbackColor);

            if (rgb is null)
                return null;

            return new A.SolidFill(rgb);
        }
        private     static  A.PatternFill               buildPatternFill        (PatternDefinition patternDefinition)                                               
        {
            if (patternDefinition is null) 
                return null;

            A.PatternFill patternFill = new () 
            { 
                Preset          = mapPatternPreset(patternDefinition.Preset),
                ForegroundColor = new A.ForegroundColor(toHexFill(patternDefinition.ForegroundColor, "000000")),
                BackgroundColor = new A.BackgroundColor(toHexFill(patternDefinition.BackgroundColor, "FFFFFF")),
            };

            return patternFill;
        }
        private     static  A.GradientFill              buildGradientFill       (GradientDefinition gradientDefinition)                                             
        {
            if (gradientDefinition is null)
                return null;
            bool                         scaled           = gradientDefinition.Scaled ?? true;
            A.GradientFill               gradientFill     = new () { RotateWithShape = new BooleanValue(!scaled) };
            A.GradientStopList           gradientStopList = new ();
            List<GradientStopDefinition> gradientStops    = gradientDefinition.Stops;

            if (gradientStops.NullOrEmpty())
            {
                gradientStops = [];

                gradientStops.Add(new GradientStopDefinition() 
                {
                    Position = 0,
                    Color    = new ColorA("#000000", null),
                });

                gradientStops.Add(new GradientStopDefinition() 
                {
                    Position = 100,
                    Color    = new ColorA("#ffffff", null),
                });
            }

            foreach (GradientStopDefinition stopDefinition in gradientStops)
            {
                A.GradientStop gradientStop = new () 
                { 
                    Position         = stopDefinition.Position * 1000, // 0..100000
                    RgbColorModelHex = toHexFill(stopDefinition.Color, "000000"),
                }; 

                gradientStopList.Append(gradientStop);
            }

            gradientFill.Append(gradientStopList);
            gradientFill.Append(new A.LinearGradientFill() 
            { 
                Angle  = new Int32Value((gradientDefinition.Angle ?? 45) * 60000), 
                Scaled = scaled
            });

            return gradientFill;
        }
        private     static  void                        buildCategoryAndValues  (SeriesDefinition s, out CategoryAxisData catAxisData, out Values values)           
        {
            catAxisData = new CategoryAxisData();
            values      = new Values();

            var stringLiteral = new StringLiteral();
            var numberLiteral = new NumberLiteral();

            stringLiteral.Append(new PointCount() { Val = (uint)s.Points.Count });

            for (int i = 0; i < s.Points.Count; i++)
            {
                stringLiteral.Append(new StringPoint()
                {
                    Index        = (uint)i,
                    NumericValue = new NumericValue(s.Points[i].Category)
                });
                numberLiteral.Append(new NumericPoint()
                {
                    Index        = (uint)i,
                    NumericValue = new NumericValue(s.Points[i].Value?.ToString(CultureInfo.InvariantCulture))
                });
            }

            catAxisData.Append(stringLiteral);
            values.Append(numberLiteral);
        }
        private     static  C.DataPoint                 buildDataPoint          (int index, PointDefinition pointDefinition, bool useMarker)                        
        {
            C.DataPoint dpt = new (new C.Index() { Val = (uint)index });

            if (useMarker)
            {
                C.Marker marker = new();

                if (pointDefinition.MarkerType is not null)
                    marker.Symbol = new C.Symbol() { Val = mapMarker(pointDefinition.MarkerType.Value) };

                if (pointDefinition.MarkerSize is not null)
                    marker.Size = new C.Size { Val = new ByteValue((byte)clamp(pointDefinition.MarkerSize.Value, 2, 72)) };

                if (pointDefinition.PointFormat is FormatDefinition fd)
                    applyFormat(marker, fd);

                dpt.Append(marker);
            }
            else
            {
                if (pointDefinition.PointFormat is FormatDefinition fd)
                    applyFormat(dpt, fd);
            }

            return dpt;
        }

        private     static  void                        ensureMarker            (OpenXmlCompositeElement series, ChartDefinition chartDefinition, SeriesDefinition seriesDefinition)    
        {
            if (series is null || chartDefinition is null || seriesDefinition is null)
                return;

            C.Marker marker = getOrAddMarker(series);

            if (marker is null)
                return;

            MarkerType? markerType = seriesDefinition.Marker?.Type ?? chartDefinition.Marker?.Type;
            int?        markerSize = seriesDefinition.Marker?.Size ?? chartDefinition.Marker?.Size;

            if (markerType is not null)
                marker.Symbol = new C.Symbol(){ Val = mapMarker(markerType.Value) };

            if (markerSize is not null)
                marker.Size = new C.Size { Val = new ByteValue((byte)clamp(markerSize.Value, 2, 72)) };

            FormatDefinition markerFormat = seriesDefinition.Marker?.Format ?? chartDefinition.Marker?.Format;

            if (markerFormat is not null)
            {
                applyFormat(marker, seriesDefinition.Marker?.Format, chartDefinition.Marker?.Format);
            }
            else
            {
                ColorA seriesSolidColor = seriesDefinition.ColorA;

                if (!string.IsNullOrWhiteSpace(seriesSolidColor?.Color))
                {
                    string seriesColor          = seriesSolidColor.Color;
                    bool   applyMarkerFillAlpha = chartDefinition.Type == ChartType.Scatter && chartDefinition.ScatterStyle == Internal.ScatterStyle.Marker;

                    ColorA markerFillColor = applyMarkerFillAlpha ? seriesSolidColor : new(seriesColor, 0);

                    C.ChartShapeProperties mSpPr = getOrAddChartShapeProps(marker);

                    // ---- Fill (içi) ----
                    mSpPr.RemoveAllChildren<A.NoFill>();
                    mSpPr.RemoveAllChildren<A.SolidFill>();     // sadece marker içi fill'i temizleyecek (ilk seviye)
                    mSpPr.Append(new A.SolidFill(toHexFill(markerFillColor)));

                    // ---- Outline (kenarlık) ----
                    A.Outline outline = mSpPr.GetFirstChild<A.Outline>();

                    if (outline is null)
                    {
                        outline = new A.Outline();
                        mSpPr.Append(outline);
                    }
                    else
                    {
                        outline.RemoveAllChildren<A.NoFill>();
                        outline.RemoveAllChildren<A.SolidFill>();
                    }

                    outline.Append(new A.SolidFill(toHexFill(seriesColor)));
                }
            }
        }
        private     static  void                        ensureNoMarker          (OpenXmlCompositeElement series)                                                    
        {
            if (series is null) 
                return;

            // Seri seviyesi: marker yok
            C.Marker seriesMarker = series.GetFirstChild<C.Marker>();

            if (seriesMarker is null)
            {
                seriesMarker = new C.Marker();

                // spPr varsa hemen sonrasına, yoksa sona eklemek güvenli
                C.ChartShapeProperties spPr = series.GetFirstChild<C.ChartShapeProperties>();

                if (spPr is null) 
                    series.Append(seriesMarker);
                else 
                    series.InsertAfter(seriesMarker, spPr);
            }

            seriesMarker.Symbol = new C.Symbol() { Val = C.MarkerStyleValues.None };

            // Nokta seviyesi: varsa marker'ları kaldır
            foreach (C.DataPoint dpt in series.Elements<C.DataPoint>())
                dpt.GetFirstChild<C.Marker>()?.Remove();
        }
        private     static  void                        ensureShapeLineVisible  (OpenXmlCompositeElement node)                                                      
        {
            if (node is null)
                return;

            C.ChartShapeProperties spPr = node.GetFirstChild<C.ChartShapeProperties>();

            if (spPr?.GetFirstChild<A.Outline>()?.GetFirstChild<A.NoFill>() is A.NoFill nf)
                nf.Remove();
        }
        private     static  void                        ensureShapeLineHidden   (OpenXmlCompositeElement node)                                                      
        {
            if (node is null)
                return;

            C.ChartShapeProperties  spPr    = getOrAddChartShapeProps(node);
            A.Outline               outline = getOrAddOutline(spPr);

            if (outline.GetFirstChild<A.NoFill>() is null)
                outline.Append(new A.NoFill());
        }

        private     static  C.ChartShapeProperties      getOrAddChartShapeProps (OpenXmlCompositeElement node)                                                      
        {
            return getOrAdd<C.ChartShapeProperties>(node);
        } 
        private     static  C.TextProperties            getOrAddTextProperties  (OpenXmlCompositeElement node)                                                      
        {
            return getOrAdd<C.TextProperties>(node);
        } 
        private     static  A.Outline                   getOrAddOutline         (OpenXmlCompositeElement node)                                                      
        {
            return getOrAdd<A.Outline>(node);
        } 
        private     static  C.Marker                    getOrAddMarker          (OpenXmlCompositeElement node)                                                      
        {
            return getOrAdd<C.Marker>(node);
        }
        private     static  T                           getOrAdd<T>             (OpenXmlCompositeElement node)                                                      where T : OpenXmlCompositeElement, new()
        {
            if (node is null)
                return null;

            T elem = node.GetFirstChild<T>();

            if (elem is null)
            {
                elem = new T();
                node.Append(elem);
            }

            return elem;
        }

        private     static  uint                        getSafeId               ()                                                                                  
        {
            byte[] guidBytes = Guid.NewGuid().ToByteArray();
            return BitConverter.ToUInt32(guidBytes, 0) & 0x7FFFFFFF; // Pozitif uint
        }
        private     static  A.RgbColorModelHex          toHexFill               (string hexOrNamed, string fallback = null, int? transparency = null)               
        {
            string hexColor = Converter.ToHexColorCode(string.IsNullOrWhiteSpace(hexOrNamed) ? fallback : hexOrNamed);

            if (string.IsNullOrWhiteSpace(hexColor))
                return null;

            A.RgbColorModelHex hex = new A.RgbColorModelHex() { Val = hexColor };

            if (transparency is not null)
                hex.Append(new A.Alpha { Val = transparencyToAlpha(transparency.Value) });

            return hex;
        }
        private     static  A.RgbColorModelHex          toHexFill               (ColorA colorA, string fallback = null)                                             
        {
            return toHexFill(colorA?.Color, fallback, colorA?.Transparency);
        }
        private     static  int                         transparencyToAlpha     (int transparencyPercent)                                                           
        {
            int t        = clamp(transparencyPercent, 0, 100);  // % şeffaflık
            int alphaVal = (int)Math.Round((100 - t) * 1000.0); // %opaklık → 0..100000

            return clamp(alphaVal, 0, 100000);            
        }
        private     static  int                         clamp                   (int v, int min, int max)                                                           
        {
            if (v < min) return min;
            if (v > max) return max;

            return v;
        }

        private     static  GroupingValues              mapGrouping             (GroupingType value)                                                                
        {
            // GroupingType.Clustered not supported here !

            return value switch
            {
                GroupingType.Stacked        => GroupingValues.Stacked       ,
                GroupingType.PercentStacked => GroupingValues.PercentStacked,
                _                           => GroupingValues.Standard      ,
            };
        }
        private     static  MarkerStyleValues           mapMarker               (MarkerType value)                                                                  
        {
            return value switch
            {
                MarkerType.Circle   => MarkerStyleValues.Circle  ,
                MarkerType.Square   => MarkerStyleValues.Square  ,
                MarkerType.Diamond  => MarkerStyleValues.Diamond ,
                MarkerType.Triangle => MarkerStyleValues.Triangle,
                MarkerType.X        => MarkerStyleValues.X       ,
                _                    => MarkerStyleValues.Circle  ,
            };
        }
        private     static  DisplayBlanksAsValues       mapBlanksDisplayedAs    (BlanksDisplayedAs value)                                                           
        {
            return value switch
            {
                BlanksDisplayedAs.Gap  => DisplayBlanksAsValues.Gap,
                BlanksDisplayedAs.Zero => DisplayBlanksAsValues.Zero,
                BlanksDisplayedAs.Span => DisplayBlanksAsValues.Span,
                _                      => DisplayBlanksAsValues.Gap
            };
        }
        private     static  BarGroupingValues           mapBarGrouping          (GroupingType value)                                                                
        {
            return value switch
            {
                GroupingType.Clustered      => BarGroupingValues.Clustered,
                GroupingType.Stacked        => BarGroupingValues.Stacked,
                GroupingType.PercentStacked => BarGroupingValues.PercentStacked,
                GroupingType.Standard       => BarGroupingValues.Standard,
                _                           => BarGroupingValues.Clustered
            };
        }
        private     static  LegendPositionValues        mapLegendPosition       (Internal.LegendPosition value)                                                     
        {
            return value switch
            {
                Internal.LegendPosition.Bottom   => LegendPositionValues.Bottom  ,
                Internal.LegendPosition.Left     => LegendPositionValues.Left    ,
                Internal.LegendPosition.Right    => LegendPositionValues.Right   ,
                Internal.LegendPosition.Top      => LegendPositionValues.Top     ,
                Internal.LegendPosition.TopRight => LegendPositionValues.TopRight,
                _                                => LegendPositionValues.Right   ,
            };
        }
        private     static  DataLabelPositionValues?    mapDataLabelPosition    (Internal.DataLabelPosition? value, ChartDefinition chartDefinition)                
        {
            if (chartDefinition is null || chartDefinition.ThreeDView is not null || (chartDefinition.Type == ChartType.Bubble && chartDefinition.Bubble3D == true))
                return null;

            ChartType chartType = chartDefinition.Type;

            switch (chartType)
            {
                case ChartType.Bar     :
                case ChartType.Column  :
                    return value switch
                    {
                        Internal.DataLabelPosition.Center      => DataLabelPositionValues.Center    ,
                        Internal.DataLabelPosition.InsideEnd   => DataLabelPositionValues.InsideEnd ,
                        Internal.DataLabelPosition.InsideBase  => DataLabelPositionValues.InsideBase,
                        Internal.DataLabelPosition.OutsideEnd  => DataLabelPositionValues.OutsideEnd,
                        _                                      => DataLabelPositionValues.OutsideEnd,
                    };
                case ChartType.Line    :
                case ChartType.Scatter :
                case ChartType.Bubble  :
                    return value switch
                    {
                        Internal.DataLabelPosition.Center      => DataLabelPositionValues.Center,
                        Internal.DataLabelPosition.Left        => DataLabelPositionValues.Left  ,
                        Internal.DataLabelPosition.Right       => DataLabelPositionValues.Right ,
                        Internal.DataLabelPosition.Top         => DataLabelPositionValues.Top   ,
                        Internal.DataLabelPosition.Bottom      => DataLabelPositionValues.Bottom,
                        _                                      => DataLabelPositionValues.Top   ,
                    };
                case ChartType.Area    :
                case ChartType.Radar   :
                case ChartType.Doughnut:
                    return null;
                case ChartType.Pie     :
                    return value switch
                    {
                        Internal.DataLabelPosition.Center      => DataLabelPositionValues.Center    ,
                        Internal.DataLabelPosition.InsideEnd   => DataLabelPositionValues.InsideEnd ,
                        Internal.DataLabelPosition.OutsideEnd  => DataLabelPositionValues.OutsideEnd,
                        Internal.DataLabelPosition.BestFit     => DataLabelPositionValues.BestFit   ,
                        _                                      => DataLabelPositionValues.BestFit   ,
                    };
                default                :
                    return null;
            }
        }
        private     static  ShapeValues                 mapShapeValues          (ThreeDShape? value)                                                                
        {
            if (value is null)
                return ShapeValues.Box;

            return value.Value switch
            {
                ThreeDShape.ConeToMax        => ShapeValues.ConeToMax       ,
                ThreeDShape.Cone             => ShapeValues.Cone            ,
                ThreeDShape.Cylinder         => ShapeValues.Cylinder        ,
                ThreeDShape.Box              => ShapeValues.Box             ,
                ThreeDShape.Pyramid          => ShapeValues.Pyramid         ,
                ThreeDShape.PyramidToMaximum => ShapeValues.PyramidToMaximum,
                _                            => ShapeValues.Box             ,
            };
        }
        private     static  A.PresetPatternValues       mapPatternPreset        (PatternFillPreset value)                                                           
        {
            return value switch 
            {
                PatternFillPreset.Percent5                => A.PresetPatternValues.Percent5              ,
                PatternFillPreset.Percent10               => A.PresetPatternValues.Percent10             ,
                PatternFillPreset.Percent20               => A.PresetPatternValues.Percent20             ,
                PatternFillPreset.Percent25               => A.PresetPatternValues.Percent25             ,
                PatternFillPreset.Percent30               => A.PresetPatternValues.Percent30             ,
                PatternFillPreset.Percent40               => A.PresetPatternValues.Percent40             ,
                PatternFillPreset.Percent50               => A.PresetPatternValues.Percent50             ,
                PatternFillPreset.Percent60               => A.PresetPatternValues.Percent60             ,
                PatternFillPreset.Percent70               => A.PresetPatternValues.Percent70             ,
                PatternFillPreset.Percent75               => A.PresetPatternValues.Percent75             ,
                PatternFillPreset.Percent80               => A.PresetPatternValues.Percent80             ,
                PatternFillPreset.Percent90               => A.PresetPatternValues.Percent90             ,
                PatternFillPreset.Horizontal              => A.PresetPatternValues.Horizontal            ,
                PatternFillPreset.Vertical                => A.PresetPatternValues.Vertical              ,
                PatternFillPreset.LightHorizontal         => A.PresetPatternValues.LightHorizontal       ,
                PatternFillPreset.LightVertical           => A.PresetPatternValues.LightVertical         ,
                PatternFillPreset.DarkHorizontal          => A.PresetPatternValues.DarkHorizontal        ,
                PatternFillPreset.DarkVertical            => A.PresetPatternValues.DarkVertical          ,
                PatternFillPreset.NarrowHorizontal        => A.PresetPatternValues.NarrowHorizontal      ,
                PatternFillPreset.NarrowVertical          => A.PresetPatternValues.NarrowVertical        ,
                PatternFillPreset.DashedHorizontal        => A.PresetPatternValues.DashedHorizontal      ,
                PatternFillPreset.DashedVertical          => A.PresetPatternValues.DashedVertical        ,
                PatternFillPreset.Cross                   => A.PresetPatternValues.Cross                 ,
                PatternFillPreset.DownwardDiagonal        => A.PresetPatternValues.DownwardDiagonal      ,
                PatternFillPreset.UpwardDiagonal          => A.PresetPatternValues.UpwardDiagonal        ,
                PatternFillPreset.LightDownwardDiagonal   => A.PresetPatternValues.LightDownwardDiagonal ,
                PatternFillPreset.LightUpwardDiagonal     => A.PresetPatternValues.LightUpwardDiagonal   ,
                PatternFillPreset.DarkDownwardDiagonal    => A.PresetPatternValues.DarkDownwardDiagonal  ,
                PatternFillPreset.DarkUpwardDiagonal      => A.PresetPatternValues.DarkUpwardDiagonal    ,
                PatternFillPreset.WideDownwardDiagonal    => A.PresetPatternValues.WideDownwardDiagonal  ,
                PatternFillPreset.WideUpwardDiagonal      => A.PresetPatternValues.WideUpwardDiagonal    ,
                PatternFillPreset.DashedDownwardDiagonal  => A.PresetPatternValues.DashedDownwardDiagonal,
                PatternFillPreset.DashedUpwardDiagonal    => A.PresetPatternValues.DashedUpwardDiagonal  ,
                PatternFillPreset.DiagonalCross           => A.PresetPatternValues.DiagonalCross         ,
                PatternFillPreset.SmallCheck              => A.PresetPatternValues.SmallCheck            ,
                PatternFillPreset.LargeCheck              => A.PresetPatternValues.LargeCheck            ,
                PatternFillPreset.SmallGrid               => A.PresetPatternValues.SmallGrid             ,
                PatternFillPreset.LargeGrid               => A.PresetPatternValues.LargeGrid             ,
                PatternFillPreset.DotGrid                 => A.PresetPatternValues.DotGrid               ,
                PatternFillPreset.SmallConfetti           => A.PresetPatternValues.SmallConfetti         ,
                PatternFillPreset.LargeConfetti           => A.PresetPatternValues.LargeConfetti         ,
                PatternFillPreset.HorizontalBrick         => A.PresetPatternValues.HorizontalBrick       ,
                PatternFillPreset.DiagonalBrick           => A.PresetPatternValues.DiagonalBrick         ,
                PatternFillPreset.SolidDiamond            => A.PresetPatternValues.SolidDiamond          ,
                PatternFillPreset.OpenDiamond             => A.PresetPatternValues.OpenDiamond           ,
                PatternFillPreset.DottedDiamond           => A.PresetPatternValues.DottedDiamond         ,
                PatternFillPreset.Plaid                   => A.PresetPatternValues.Plaid                 ,
                PatternFillPreset.Sphere                  => A.PresetPatternValues.Sphere                ,
                PatternFillPreset.Weave                   => A.PresetPatternValues.Weave                 ,
                PatternFillPreset.Divot                   => A.PresetPatternValues.Divot                 ,
                PatternFillPreset.Shingle                 => A.PresetPatternValues.Shingle               ,
                PatternFillPreset.Wave                    => A.PresetPatternValues.Wave                  ,
                PatternFillPreset.Trellis                 => A.PresetPatternValues.Trellis               ,
                PatternFillPreset.ZigZag                  => A.PresetPatternValues.ZigZag                ,
                _                                         => A.PresetPatternValues.DotGrid
            };
        }
        private     static  C.ScatterStyleValues        mapScatterStyle         (Internal.ScatterStyle value)                                                       
        {
            // GroupingType.Clustered not supported here !

            return value switch
            {
                Internal.ScatterStyle.Line         => C.ScatterStyleValues.Line        ,
                Internal.ScatterStyle.LineMarker   => C.ScatterStyleValues.LineMarker  ,
                Internal.ScatterStyle.Marker       => C.ScatterStyleValues.Marker      ,
                Internal.ScatterStyle.Smooth       => C.ScatterStyleValues.Smooth      ,
                Internal.ScatterStyle.SmoothMarker => C.ScatterStyleValues.SmoothMarker,
                _                                  => C.ScatterStyleValues.Marker      ,
            };
        }
        private     static  C.RadarStyleValues          mapRadarStyle           (Internal.RadarStyle value)                                                         
        {
            return value switch
            {
                Internal.RadarStyle.Standard => C.RadarStyleValues.Standard,
                Internal.RadarStyle.Marker   => C.RadarStyleValues.Marker,
                Internal.RadarStyle.Filled   => C.RadarStyleValues.Filled,
                _                            => C.RadarStyleValues.Standard,
            };
        }
        private     static  A.PresetLineDashValues      mapDash                 (DashPreset value)                                                                  
        {
            return value switch
            {
                DashPreset.Solid            => A.PresetLineDashValues.Solid,
                DashPreset.Dot              => A.PresetLineDashValues.Dot,
                DashPreset.Dash             => A.PresetLineDashValues.Dash,
                DashPreset.LargeDash        => A.PresetLineDashValues.LargeDash,
                DashPreset.DashDot          => A.PresetLineDashValues.DashDot,
                DashPreset.LargeDashDot     => A.PresetLineDashValues.LargeDashDot,
                DashPreset.LargeDashDotDot  => A.PresetLineDashValues.LargeDashDotDot,
                DashPreset.SystemDash       => A.PresetLineDashValues.SystemDash,
                DashPreset.SystemDot        => A.PresetLineDashValues.SystemDot,
                DashPreset.SystemDashDot    => A.PresetLineDashValues.SystemDashDot,
                DashPreset.SystemDashDotDot => A.PresetLineDashValues.SystemDashDotDot,
                _                           => A.PresetLineDashValues.Solid
            };
        }
        private     static  A.CompoundLineValues        mapCompound             (CompoundLinePreset value)                                                          
        {
            return value switch
            {
                CompoundLinePreset.Single    => A.CompoundLineValues.Single,
                CompoundLinePreset.Double    => A.CompoundLineValues.Double,
                CompoundLinePreset.ThickThin => A.CompoundLineValues.ThickThin,
                CompoundLinePreset.ThinThick => A.CompoundLineValues.ThinThick,
                CompoundLinePreset.Triple    => A.CompoundLineValues.Triple,
                _                            => A.CompoundLineValues.Single
            };
        }
        private     static  A.LineCapValues             mapLineCap              (LineCapPreset value)                                                               
        {
            return value switch
            {
                LineCapPreset.Flat   => A.LineCapValues.Flat,
                LineCapPreset.Round  => A.LineCapValues.Round,
                LineCapPreset.Square => A.LineCapValues.Square,
                _                    => A.LineCapValues.Flat
            };
        }
        private     static  A.LineEndValues             mapLineEndType          (LineEndPreset value)                                                               
        {
            return value switch
            {
                LineEndPreset.None     => A.LineEndValues.None,
                LineEndPreset.Triangle => A.LineEndValues.Triangle,
                LineEndPreset.Stealth  => A.LineEndValues.Stealth,
                LineEndPreset.Diamond  => A.LineEndValues.Diamond,
                LineEndPreset.Oval     => A.LineEndValues.Oval,
                LineEndPreset.Arrow    => A.LineEndValues.Arrow,
                _                      => A.LineEndValues.None
            };
        }
        private     static  A.LineEndWidthValues        mapLineEndWidth         (LineEndWidthPreset value)                                                          
        {
            return value switch
            {
                LineEndWidthPreset.Small  => A.LineEndWidthValues.Small,
                LineEndWidthPreset.Medium => A.LineEndWidthValues.Medium,
                LineEndWidthPreset.Large  => A.LineEndWidthValues.Large,
                _                         => A.LineEndWidthValues.Medium
            };
        }
        private     static  A.LineEndLengthValues       mapLineEndLength        (LineEndLengthPreset value)                                                         
        {
            return value switch
            {
                LineEndLengthPreset.Small  => A.LineEndLengthValues.Small,
                LineEndLengthPreset.Medium => A.LineEndLengthValues.Medium,
                LineEndLengthPreset.Large  => A.LineEndLengthValues.Large,
                _                          => A.LineEndLengthValues.Medium
            };
        }
        private     static  A.BevelPresetValues         mapBevel                (BevelPreset value)                                                                 
        {
            return value switch 
            {
                BevelPreset.RelaxedInset => A.BevelPresetValues.RelaxedInset,
                BevelPreset.Circle       => A.BevelPresetValues.Circle,
                BevelPreset.Slope        => A.BevelPresetValues.Slope,
                BevelPreset.Cross        => A.BevelPresetValues.Cross,
                BevelPreset.Angle        => A.BevelPresetValues.Angle,
                BevelPreset.SoftRound    => A.BevelPresetValues.SoftRound,
                BevelPreset.Convex       => A.BevelPresetValues.Convex,
                BevelPreset.CoolSlant    => A.BevelPresetValues.CoolSlant,
                BevelPreset.Divot        => A.BevelPresetValues.Divot,
                BevelPreset.Riblet       => A.BevelPresetValues.Riblet,
                BevelPreset.HardEdge     => A.BevelPresetValues.HardEdge,
                BevelPreset.ArtDeco      => A.BevelPresetValues.ArtDeco,
                _                        => A.BevelPresetValues.RelaxedInset
            };
        }
        private     static  A.PresetMaterialTypeValues  mapMaterial             (MaterialPreset value)                                                              
        {
            return value switch 
            {
                MaterialPreset.LegacyMatte        => A.PresetMaterialTypeValues.LegacyMatte,
                MaterialPreset.LegacyPlastic      => A.PresetMaterialTypeValues.LegacyPlastic,
                MaterialPreset.LegacyMetal        => A.PresetMaterialTypeValues.LegacyMetal,
                MaterialPreset.LegacyWireframe    => A.PresetMaterialTypeValues.LegacyWireframe,
                MaterialPreset.Matte              => A.PresetMaterialTypeValues.Matte,
                MaterialPreset.Plastic            => A.PresetMaterialTypeValues.Plastic,
                MaterialPreset.Metal              => A.PresetMaterialTypeValues.Metal,
                MaterialPreset.WarmMatte          => A.PresetMaterialTypeValues.WarmMatte,
                MaterialPreset.TranslucentPowder  => A.PresetMaterialTypeValues.TranslucentPowder,
                MaterialPreset.Powder             => A.PresetMaterialTypeValues.Powder,
                MaterialPreset.DarkEdge           => A.PresetMaterialTypeValues.DarkEdge,
                MaterialPreset.SoftEdge           => A.PresetMaterialTypeValues.SoftEdge,
                MaterialPreset.Clear              => A.PresetMaterialTypeValues.Clear,
                MaterialPreset.Flat               => A.PresetMaterialTypeValues.Flat,
                MaterialPreset.SoftMetal          => A.PresetMaterialTypeValues.SoftMetal,
                _                                 => A.PresetMaterialTypeValues.Matte
            };
        }
        private     static  A.LightRigDirectionValues   mapLightDir             (LightingDirection value)                                                           
        {
            return value switch 
            {
                LightingDirection.TopLeft     => A.LightRigDirectionValues.TopLeft,
                LightingDirection.Top         => A.LightRigDirectionValues.Top,
                LightingDirection.TopRight    => A.LightRigDirectionValues.TopRight,
                LightingDirection.Left        => A.LightRigDirectionValues.Left,
                LightingDirection.Right       => A.LightRigDirectionValues.Right,
                LightingDirection.BottomLeft  => A.LightRigDirectionValues.BottomLeft,
                LightingDirection.Bottom      => A.LightRigDirectionValues.Bottom,
                LightingDirection.BottomRight => A.LightRigDirectionValues.BottomRight,
                _                             => A.LightRigDirectionValues.Top
            };
        }
        private     static  A.LightRigValues            mapLightPreset          (LightingPreset value)                                                              
        {
            return value switch 
            {
                LightingPreset.LegacyFlat1   => A.LightRigValues.LegacyFlat1,
                LightingPreset.LegacyFlat2   => A.LightRigValues.LegacyFlat2,
                LightingPreset.LegacyFlat3   => A.LightRigValues.LegacyFlat3,
                LightingPreset.LegacyFlat4   => A.LightRigValues.LegacyFlat4,
                LightingPreset.LegacyNormal1 => A.LightRigValues.LegacyNormal1,
                LightingPreset.LegacyNormal2 => A.LightRigValues.LegacyNormal2,
                LightingPreset.LegacyNormal3 => A.LightRigValues.LegacyNormal3,
                LightingPreset.LegacyNormal4 => A.LightRigValues.LegacyNormal4,
                LightingPreset.LegacyHarsh1  => A.LightRigValues.LegacyHarsh1,
                LightingPreset.LegacyHarsh2  => A.LightRigValues.LegacyHarsh2,
                LightingPreset.LegacyHarsh3  => A.LightRigValues.LegacyHarsh3,
                LightingPreset.LegacyHarsh4  => A.LightRigValues.LegacyHarsh4,
                LightingPreset.ThreePoints   => A.LightRigValues.ThreePoints,
                LightingPreset.Balanced      => A.LightRigValues.Balanced,
                LightingPreset.Soft          => A.LightRigValues.Soft,
                LightingPreset.Harsh         => A.LightRigValues.Harsh,
                LightingPreset.Flood         => A.LightRigValues.Flood,
                LightingPreset.Contrasting   => A.LightRigValues.Contrasting,
                LightingPreset.Morning       => A.LightRigValues.Morning,
                LightingPreset.Sunrise       => A.LightRigValues.Sunrise,
                LightingPreset.Sunset        => A.LightRigValues.Sunset,
                LightingPreset.Chilly        => A.LightRigValues.Chilly,
                LightingPreset.Freezing      => A.LightRigValues.Freezing,
                LightingPreset.Flat          => A.LightRigValues.Flat,
                LightingPreset.TwoPoints     => A.LightRigValues.TwoPoints,
                LightingPreset.Glow          => A.LightRigValues.Glow,
                LightingPreset.BrightRoom    => A.LightRigValues.BrightRoom,
                _                            => A.LightRigValues.Soft
            };
        }
        private     static  A.PresetShadowValues        mapShadowPreset         (ShadowPreset? value)                                                               
        {
            if (value is null)
                return A.PresetShadowValues.FrontBottomShadow;

            return value.Value switch 
            {
                ShadowPreset.TopLeftDropShadow               => A.PresetShadowValues.TopLeftDropShadow              ,
                ShadowPreset.TopRightDropShadow              => A.PresetShadowValues.TopRightDropShadow             ,
                ShadowPreset.BackLeftPerspectiveShadow       => A.PresetShadowValues.BackLeftPerspectiveShadow      ,
                ShadowPreset.BackRightPerspectiveShadow      => A.PresetShadowValues.BackRightPerspectiveShadow     ,
                ShadowPreset.BottomLeftDropShadow            => A.PresetShadowValues.BottomLeftDropShadow           ,
                ShadowPreset.BottomRightDropShadow           => A.PresetShadowValues.BottomRightDropShadow          ,
                ShadowPreset.FrontLeftPerspectiveShadow      => A.PresetShadowValues.FrontLeftPerspectiveShadow     ,
                ShadowPreset.FrontRightPerspectiveShadow     => A.PresetShadowValues.FrontRightPerspectiveShadow    ,
                ShadowPreset.TopLeftSmallDropShadow          => A.PresetShadowValues.TopLeftSmallDropShadow         ,
                ShadowPreset.TopLeftLargeDropShadow          => A.PresetShadowValues.TopLeftLargeDropShadow         ,
                ShadowPreset.BackLeftLongPerspectiveShadow   => A.PresetShadowValues.BackLeftLongPerspectiveShadow  ,
                ShadowPreset.BackRightLongPerspectiveShadow  => A.PresetShadowValues.BackRightLongPerspectiveShadow ,
                ShadowPreset.TopLeftDoubleDropShadow         => A.PresetShadowValues.TopLeftDoubleDropShadow        ,
                ShadowPreset.BottomRightSmallDropShadow      => A.PresetShadowValues.BottomRightSmallDropShadow     ,
                ShadowPreset.FrontLeftLongPerspectiveShadow  => A.PresetShadowValues.FrontLeftLongPerspectiveShadow ,
                ShadowPreset.FrontRightLongPerspectiveShadow => A.PresetShadowValues.FrontRightLongPerspectiveShadow,
                ShadowPreset.ThreeDimensionalOuterBoxShadow  => A.PresetShadowValues.ThreeDimensionalOuterBoxShadow ,
                ShadowPreset.ThreeDimensionalInnerBoxShadow  => A.PresetShadowValues.ThreeDimensionalInnerBoxShadow ,
                ShadowPreset.BackCenterPerspectiveShadow     => A.PresetShadowValues.BackCenterPerspectiveShadow    ,
                ShadowPreset.FrontBottomShadow               => A.PresetShadowValues.FrontBottomShadow              ,
                _                                            => A.PresetShadowValues.FrontBottomShadow              ,
            };
        }
        private     static  A.TextAlignmentTypeValues   mapTextAlign            (HorizontalAlign value)                                                             
        {
            return value switch 
            {
                HorizontalAlign.Left    => A.TextAlignmentTypeValues.Left,
                HorizontalAlign.Center  => A.TextAlignmentTypeValues.Center,
                HorizontalAlign.Right   => A.TextAlignmentTypeValues.Right,
                HorizontalAlign.Justify => A.TextAlignmentTypeValues.Justified,
                _                       => A.TextAlignmentTypeValues.Left
            };
        }
        private     static  A.TextAnchoringTypeValues   mapTextAnchor           (VerticalAlign value)                                                               
        {
            return value switch 
            {
                VerticalAlign.Top      => A.TextAnchoringTypeValues.Top,
                VerticalAlign.Center   => A.TextAnchoringTypeValues.Center,
                VerticalAlign.Bottom   => A.TextAnchoringTypeValues.Bottom,
                _                      => A.TextAnchoringTypeValues.Center
            };
        }
        private     static  A.TextWrappingValues        mapTextWrap             (TextWrapPreset value)                                                              
        {
            return value switch 
            {
                TextWrapPreset.None   => A.TextWrappingValues.None,
                TextWrapPreset.Square => A.TextWrappingValues.Square,
                _                     => A.TextWrappingValues.Square
            };
        }
        private     static  A.TextUnderlineValues       mapUnderline            (TextUnderlineStyle value)                                                          
        {
            return value switch 
            {
                TextUnderlineStyle.None   => A.TextUnderlineValues.None,
                TextUnderlineStyle.Single => A.TextUnderlineValues.Single,
                TextUnderlineStyle.Double => A.TextUnderlineValues.Double,
                _                         => A.TextUnderlineValues.None
            };
        }
        private     static  A.TextStrikeValues          mapStrike               (TextStrikeStyle value)                                                             
        {
            return value switch 
            {
                TextStrikeStyle.None   => A.TextStrikeValues.NoStrike,
                TextStrikeStyle.Single => A.TextStrikeValues.SingleStrike,
                TextStrikeStyle.Double => A.TextStrikeValues.DoubleStrike,
                _                      => A.TextStrikeValues.NoStrike
            };
        }
        private     static  A.TextCapsValues            mapCaps                 (TextCapsStyle value)                                                               
        {
            return value switch 
            {
                TextCapsStyle.None  => A.TextCapsValues.None,
                TextCapsStyle.Small => A.TextCapsValues.Small,
                TextCapsStyle.All   => A.TextCapsValues.All,
                _                   => A.TextCapsValues.None
            };
        }
    }
}

#if false

            if (false && chartDefinition?.VaryColors != true && !string.IsNullOrWhiteSpace(seriesDefinition.Color))
            {
                bool applyMarkerFillAlpha = chartDefinition.Type == ChartType.Scatter && chartDefinition.ScatterStyle == Internal.ScatterStyle.Marker;

                C.ChartShapeProperties mSpPr = marker.GetFirstChild<C.ChartShapeProperties>();

                if (mSpPr is null)
                {
                    mSpPr = new C.ChartShapeProperties();
                    marker.Append(mSpPr);
                }

                string hex = Converter.ToHexColorCode(seriesDefinition.Color);

                // ---- Fill (içi) ----
                mSpPr.RemoveAllChildren<A.NoFill>();
                mSpPr.RemoveAllChildren<A.SolidFill>();     // sadece marker içi fill'i temizleyecek (ilk seviye)
                A.RgbColorModelHex fillHex = new A.RgbColorModelHex { Val = hex };

                if (applyMarkerFillAlpha && seriesDefinition.ColorTransparency is not null)
                    fillHex.Append(new A.Alpha { Val = transparencyToAlpha(seriesDefinition.ColorTransparency.Value) });

                mSpPr.Append(new A.SolidFill(fillHex));

                // ---- Outline (kenarlık) ----
                A.Outline outline = mSpPr.GetFirstChild<A.Outline>();

                if (outline is null)
                {
                    outline = new A.Outline();
                    mSpPr.Append(outline);
                }
                else
                {
                    outline.RemoveAllChildren<A.NoFill>();
                    outline.RemoveAllChildren<A.SolidFill>();
                }

                outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = hex }));
            }

        private     static  void                        applySeriesColor_old    (ChartDefinition chartDefinition, SeriesDefinition seriesDefinition, OpenXmlCompositeElement series, bool toOutline = false)
        {
            // Seri rengi (VaryColors != true ise sabitle)
            if (chartDefinition is null || seriesDefinition is null || series is null || chartDefinition.VaryColors == true || string.IsNullOrWhiteSpace(seriesDefinition.Color))
                return;

            C.ChartShapeProperties spPr = series.GetFirstChild<C.ChartShapeProperties>();

            if (spPr is null)
            {
                spPr = new C.ChartShapeProperties();
                series.Append(spPr);
            }

            A.RgbColorModelHex fillHex = new A.RgbColorModelHex() { Val = Converter.ToHexColorCode(seriesDefinition.Color) };

            if (toOutline)
            {
                A.Outline outline = spPr.GetFirstChild<A.Outline>();

                if (outline is null)
                {
                    outline = new A.Outline();
                    spPr.Append(outline);
                }
                else
                {
                    outline.RemoveAllChildren<A.SolidFill>();
                    outline.RemoveAllChildren<A.NoFill>();
                }

                if (seriesDefinition.ColorTransparency is not null && chartDefinition.IsLineBased)
                    fillHex.Append(new A.Alpha { Val = transparencyToAlpha(seriesDefinition.ColorTransparency.Value) });

                outline.Append(new A.SolidFill(fillHex));
            }
            else
            {
                spPr.RemoveAllChildren<A.SolidFill>();
                spPr.RemoveAllChildren<A.NoFill>();

                if (seriesDefinition.ColorTransparency is not null)
                    fillHex.Append(new A.Alpha { Val = transparencyToAlpha(seriesDefinition.ColorTransparency.Value) });

                spPr.Append(new A.SolidFill(fillHex));
            }
        }
        private     static  A.EffectList                getOrAddEffectList      (OpenXmlCompositeElement node)
        {
            if (node is null)
                return null;

            A.EffectList effectList = node.GetFirstChild<A.EffectList>();

            if (effectList is null)
            {
                effectList = new A.EffectList();
                node.Append(effectList);
            }

            return effectList;
        } 
        private     static  C.Floor                     ensureFloor             (OpenXmlCompositeElement plotArea)
        {
            C.Floor floor = plotArea.GetFirstChild<C.Floor>();

            if (floor is null) 
            { 
                floor = new C.Floor(); 
                plotArea.Append(floor); 
            }

            getOrAddChartShapeProps(floor);

            return floor;
        }
        private     static  C.SideWall                  ensureSideWall          (OpenXmlCompositeElement plotArea)
        {
            C.SideWall sideWall = plotArea.GetFirstChild<C.SideWall>();

            if (sideWall is null) 
            { 
                sideWall = new C.SideWall();
                plotArea.Append(sideWall); 
            }

            getOrAddChartShapeProps(sideWall);

            return sideWall;
        }
        private     static  C.BackWall                  ensureBackWall          (OpenXmlCompositeElement plotArea)
        {
            C.BackWall backWall = plotArea.GetFirstChild<C.BackWall>();

            if (backWall is null)
            {
                backWall = new C.BackWall();
                plotArea.Append(backWall);
            }

            getOrAddChartShapeProps(backWall);

            return backWall;
        }
        private     static  C.DataLabels                ensureDataLabels        (OpenXmlCompositeElement node)
        {
            C.DataLabels dataLabels = node.GetFirstChild<C.DataLabels>();

            dataLabels ??= node.AppendChild(new C.DataLabels());

            return dataLabels;
        }

        private     static  A.Outline                   buildBorder             (string color, Dimension width)
        {
            A.Outline outline = new A.Outline(new A.SolidFill(toHexFill(color, "000000")));

            if (width is not null)
                outline.Width = width.ToEmu();

            return outline;
        }
        private     static  A.SolidFill                 buildSolidFill          (string color)                                                                  
        {
            return new A.SolidFill(toHexFill(color, "dedede"));
        }
        private     static  A.PatternFill               buildPatternFill        (PatternFillPreset patternFillPreset, string foreGround, string backGround)     
        {
            A.PatternFill p = new A.PatternFill() { Preset = mapPatternPreset(patternFillPreset) };

            p.ForegroundColor = new A.ForegroundColor(toHexFill(foreGround, "000000"));
            p.BackgroundColor = new A.BackgroundColor(toHexFill(backGround, "FFFFFF"));

            return p;
        }
        private     static  A.GradientFill              buildLinearGradientFill (string color1, string color2, int? angleDegrees)                               
        {
            A.GradientFill       gradientFill       = new A.GradientFill() { RotateWithShape = new BooleanValue(true) };
            A.LinearGradientFill linearGradientFill = new A.LinearGradientFill() { Angle = new Int32Value((angleDegrees ?? 45) * 60000) };
            A.GradientStopList   gradientStopList   = new A.GradientStopList();
            A.GradientStop       gradientStop1      = new A.GradientStop(toHexFill(color1, "000000")) { Position = new Int32Value(0) };
            A.GradientStop       gradientStop2      = new A.GradientStop(toHexFill(color2, "FFFFFF")) { Position = new Int32Value(100000) }; // 0..100000

            gradientStopList.Append(gradientStop1);
            gradientStopList.Append(gradientStop2);
            gradientFill    .Append(gradientStopList);
            gradientFill    .Append(linearGradientFill);

            return gradientFill;
        }
        private     static  void                        applyTickLabelRotation  (OpenXmlCompositeElement axisNode, int? rotation)                               
        {
            if (!rotation.HasValue) 
                return;
            // OpenXML rot = derece * 60000, saat yönünün tersine
            C.TextProperties txPr = axisNode.GetFirstChild<C.TextProperties>();
            if (txPr is null) 
            {
                txPr = new C.TextProperties(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text("")))
                );
                axisNode.Append(txPr);
            }
            txPr.BodyProperties ??= new A.BodyProperties();
            txPr.BodyProperties.Rotation = rotation.Value * 60000;
        }

        private     static  T                           getWallOrFloor<T>       (bool show, IFillAndBorderContainer fillAndBorderContainer, FillStyle fillStyle, IFillAndBorderContainer defaultFillAndBorderContainer = null, int? gradientAngleOverride = null)                                                                     where T : OpenXmlCompositeElement, new()
        {
            T elem = new T();
            if (!show)
            {
                // görünmez yapmak için: NoFill + NoLine
                elem.Append(new C.ChartShapeProperties(
                    new A.NoFill(),
                    new A.Outline(new A.NoFill())
                ));
            }
            else
            {
                applyFillAndBorder(elem, fillAndBorderContainer, fillStyle, defaultFillAndBorderContainer, gradientAngleOverride);
            }
            return elem;
        }
        private     static  void                        applyFillAndBorder      (OpenXmlCompositeElement node, IFillAndBorderContainer definition, FillStyle fillStyle, IFillAndBorderContainer defaultDefinition = null, int? gradientAngleOverride = null)
        {
            if (node is null || definition is null)
                return;

            C.ChartShapeProperties shapeProperties = new();

            bool tryAddBorder      (IFillAndBorderContainer container)
            {
                if (container is null || !container.HasBorder())
                    return false;

                shapeProperties.Append(buildBorder(container.BorderColor, container.BorderWidth));

                return true;
            }
            bool tryAddGradientFill(IFillAndBorderContainer container, bool tryFallback)
            {
                if (container is null)
                    return false;

                if (container.HasGradientFill())
                {
                    shapeProperties.Append(buildLinearGradientFill(container.GradientFillColor1, container.GradientFillColor2, gradientAngleOverride ?? container.GradientAngle));

                    return true;
                }
                else if (tryFallback)
                {
                    return tryAddPatternFill(container, true);
                }
                
                return false;
            }
            bool tryAddPatternFill (IFillAndBorderContainer container, bool tryFallback)
            {
                if (container is null)
                    return false;

                if (container.HasPatternFill())
                {
                    shapeProperties.Append(buildPatternFill(container.PatternFillPreset.Value, container.PatternFillForeground, container.PatternFillBackground));

                    return true;
                }
                else if (tryFallback)
                {
                    return tryAddSolidFill(container);
                }
                    
                return false;
            }
            bool tryAddSolidFill   (IFillAndBorderContainer container)
            {
                if (container is null)
                    return false;

                if (container.HasSolidFill())
                {
                    shapeProperties.Append(buildSolidFill(container.SolidFillColor));

                    return true;
                }

                return false;
            }

            switch (fillStyle)
            {
                case FillStyle.Auto:
                    if (!tryAddGradientFill(definition, true))
                        tryAddGradientFill(defaultDefinition, true);
                    break;
                case FillStyle.Gradient:
                    if (!tryAddGradientFill(definition, false))
                        tryAddGradientFill(defaultDefinition, false);
                    break;
                case FillStyle.Pattern :
                    if (!tryAddPatternFill(definition, false))
                        tryAddPatternFill(defaultDefinition, false);
                    break;
                case FillStyle.Solid   :
                    if (!tryAddSolidFill(definition))
                        tryAddSolidFill(defaultDefinition);
                    break;
                case FillStyle.None    :
                default:
                    break;
            }

            if (!tryAddBorder(definition))
                tryAddBorder(defaultDefinition);

            if (shapeProperties.HasChildren)
                node.Append(shapeProperties);
        }

#endif
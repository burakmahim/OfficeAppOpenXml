using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using System.Xml.Linq;
using OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents;
using System.Globalization;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeAppOpenXmlLibrary.PowerPointOpenXmlComponents
{
    class ChartComponent
    {
        private static ChartDefinition GetChartDefinitionFromXml(XElement chartNode)
        {
            ChartDefinition chartDefinition = new ChartDefinition();

            //chart

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

            Nullable<BlanksDisplayedAs> displayBlanksAs = XElementAttributeGetter.AsEnum<BlanksDisplayedAs>(chartNode, "display-blanks-as");
            if (displayBlanksAs.HasValue)
                chartDefinition.BlanksDisplayedAs = displayBlanksAs.Value;

            Nullable<MarkerType> markerType = XElementAttributeGetter.AsEnum<MarkerType>(chartNode, "marker-type");
            if (markerType.HasValue)
            {
                if (chartDefinition.Marker == null)
                    chartDefinition.Marker = new MarkerDefinition();
                chartDefinition.Marker.Type = markerType.Value;
            }

            XElementAttributeGetter.AsBool(chartNode, "vary-colors", out bool varyColorsValue);
            chartDefinition.VaryColors = varyColorsValue;

            XElementAttributeGetter.AsBool(chartNode, "auto-title", out bool autoTitleValue, defaultValue: false);
            chartDefinition.AutoTitle = autoTitleValue;

            XElementAttributeGetter.AsBool(chartNode, "show-category-axis", out bool showCategoryAxisValue, defaultValue: true);
            chartDefinition.ShowCategoryAxis = showCategoryAxisValue;

            XElementAttributeGetter.AsBool(chartNode, "show-value-axis", out bool showValueAxisValue, defaultValue: true);
            chartDefinition.ShowValueAxis = showValueAxisValue;

            //// Width ve Height okuması
            //string widthStr = chartNode.Attribute("width")?.Value;
            //if (!string.IsNullOrEmpty(widthStr))
            //    chartDefinition.Width = Dimension.Parse(widthStr);

            //string heightStr = chartNode.Attribute("height")?.Value;
            //if (!string.IsNullOrEmpty(heightStr))
            //    chartDefinition.Height = Dimension.Parse(heightStr);

            XElementAttributeGetter.AsBool(chartNode, "plot-visible-only", out bool plotVisibleOnlyValue, defaultValue: false);
            chartDefinition.PlotVisibleOnly = plotVisibleOnlyValue;

            XElementAttributeGetter.AsBool(chartNode, "rounded-corners", out bool roundedCornersValue, defaultValue: false);
            chartDefinition.RoundedCorners = roundedCornersValue;

            XElementAttributeGetter.AsBool(chartNode, "show-data-labels-over-maximum", out bool showDataLabelsOverMaximumValue, defaultValue: false);
            chartDefinition.ShowDataLabelsOverMaximum = showDataLabelsOverMaximumValue;

            XElementAttributeGetter.AsBool(chartNode, "bubble-3d", out bool bubble3DValue, defaultValue: false);
            chartDefinition.Bubble3D = bubble3DValue;

            XElementAttributeGetter.AsBool(chartNode, "create-new-paragraph", out bool createNewParagraphValue);
            chartDefinition.CreateNewParagraph = createNewParagraphValue;

            if (XElementAttributeGetter.AsInt32(chartNode, "overlap", out int overlapValue))
                chartDefinition.Overlap = clamp(overlapValue, -100, 100);

            if (XElementAttributeGetter.AsInt32(chartNode, "gap-width", out int gapWidthValue))
                chartDefinition.GapWidth = clamp(gapWidthValue, 0, 500);

            if (XElementAttributeGetter.AsInt32(chartNode, "first-slice-angle", out int firstSliceAngleValue))
                chartDefinition.FirstSliceAngle = clamp(firstSliceAngleValue, 0, 360);

            if (XElementAttributeGetter.AsInt32(chartNode, "doughnut-hole-size", out int doughnutHoleSizeValue))
                chartDefinition.DoughnutHoleSize = clamp(doughnutHoleSizeValue, 10, 90);

            if (XElementAttributeGetter.AsInt32(chartNode, "bubble-scale", out int bubbleScaleValue))
                chartDefinition.BubbleScale = clamp(bubbleScaleValue, 0, 300);

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

                XElement? textFormat = categoryAxisNode.Element("text-format");
                if (textFormat != null)
                {
                    chartDefinition.CategoryAxis.TickLabelTextFormat = XElementAttributeGetter.ParseTextFormat(textFormat);
                }

                XElement? format = categoryAxisNode.Element("format");
                if (format != null)
                {
                    chartDefinition.CategoryAxis.AxisLineFormat = XElementAttributeGetter.ParseFormatNode(format);
                }

            }

            //value-axis

            XElement? valueAxisNode = chartNode.Element("value-axis");
            if (valueAxisNode != null)
            {
                chartDefinition.ValueAxis = new AxisDefinition();

                if (XElementAttributeGetter.AsDouble(valueAxisNode, "min", out double minValue))
                    chartDefinition.ValueAxis.Min = minValue;

                if (XElementAttributeGetter.AsDouble(valueAxisNode, "max", out double maxValue))
                    chartDefinition.ValueAxis.Max = maxValue;

                if (XElementAttributeGetter.AsDouble(valueAxisNode, "major-unit", out double majorUnit))
                    chartDefinition.ValueAxis.MajorUnit = majorUnit;

                XElement? format = valueAxisNode.Element("format");
                if (format != null)
                {
                    chartDefinition.ValueAxis.AxisLineFormat = XElementAttributeGetter.ParseFormatNode(format);
                }

                XElement? textFormat = valueAxisNode.Element("text-format");
                if (textFormat != null)
                {
                    chartDefinition.ValueAxis.TickLabelTextFormat = XElementAttributeGetter.ParseTextFormat(textFormat);
                }

            }

            //legend

            XElement? legendNode = chartNode.Element("legend");
            if (legendNode != null)
            {
                chartDefinition.Legend = new LegendDefinition();

                if (XElementAttributeGetter.AsBool(legendNode, "show", out bool show))
                    chartDefinition.Legend.Show = show;

                Nullable<LegendPosition> position = XElementAttributeGetter.AsEnum<LegendPosition>(legendNode, "position");
                if (position.HasValue)
                    chartDefinition.Legend.Position = position.Value;


                XElement? format = legendNode.Element("format");
                if (format != null)
                {
                    chartDefinition.Legend.BoxFormat = XElementAttributeGetter.ParseFormatNode(format);
                }

                XElement? textFormat = legendNode.Element("text-format");
                if (textFormat != null)
                {
                    chartDefinition.Legend.LegendEntryTextFormat = XElementAttributeGetter.ParseTextFormat(textFormat);
                }

            }

            //grid

            XElement? gridNode = chartNode.Element("grid");
            if (gridNode != null)
            {
                chartDefinition.Grid = new GridDefinition();

                if (XElementAttributeGetter.AsBool(gridNode, "show-horizontal", out bool showHorizontal))
                    chartDefinition.Grid.ShowHorizontal = showHorizontal;

                if (XElementAttributeGetter.AsBool(gridNode, "show-vertical", out bool showVertical))
                    chartDefinition.Grid.ShowVertical = showVertical;

                XElement? horizontalFormatNode = gridNode.Element("horizontal-format");
                if (horizontalFormatNode != null)
                {
                    chartDefinition.Grid.HorizontalFormat = new FormatDefinition();

                    XElement? lineNode = horizontalFormatNode.Element("line");
                    if (lineNode != null)
                    {
                        chartDefinition.Grid.HorizontalFormat.LineDefinition = new LineDefinition();

                        XElementAttributeGetter.AsBool(lineNode, "visible", out bool visible);
                        chartDefinition.Grid.HorizontalFormat.LineDefinition.Visible = visible;

                        Nullable<DashPreset> dashPreset = XElementAttributeGetter.AsEnum<DashPreset>(lineNode, "dash");
                        if (dashPreset.HasValue)
                            chartDefinition.Grid.HorizontalFormat.LineDefinition.DashPreset = dashPreset.Value;

                        Nullable<CompoundLinePreset> compoundLine = XElementAttributeGetter.AsEnum<CompoundLinePreset>(lineNode, "compound");
                        if (compoundLine.HasValue)
                            chartDefinition.Grid.HorizontalFormat.LineDefinition.CompoundPreset = compoundLine.Value;

                    }

                    XElement? effectsNode = horizontalFormatNode.Element("effects");
                    if (effectsNode != null)
                    {
                        chartDefinition.Grid.HorizontalFormat.EffectsDefinition = new EffectsDefinition();

                        XElement? shadowNode = effectsNode.Element("shadow");
                        if (shadowNode != null)
                        {
                            chartDefinition.Grid.HorizontalFormat.EffectsDefinition.ShadowDefinition = new ShadowDefinition();

                            Nullable<ShadowType> shadowType = XElementAttributeGetter.AsEnum<ShadowType>(shadowNode, "type");
                            if (shadowType.HasValue)
                                chartDefinition.Grid.HorizontalFormat.EffectsDefinition.ShadowDefinition.Type = shadowType.Value;

                            Nullable<ShadowPreset> shadowPreset = XElementAttributeGetter.AsEnum<ShadowPreset>(shadowNode, "preset");
                            if (shadowPreset.HasValue)
                                chartDefinition.Grid.HorizontalFormat.EffectsDefinition.ShadowDefinition.Preset = shadowPreset.Value;

                            if (XElementAttributeGetter.AsInt32(shadowNode, "angle", out int angle))
                                chartDefinition.Grid.HorizontalFormat.EffectsDefinition.ShadowDefinition.Angle = clamp(angle, 0, 360);

                            string? blurStr = shadowNode.Attribute("blur-radius")?.Value;
                            if (!string.IsNullOrEmpty(blurStr))
                                chartDefinition.Grid.HorizontalFormat.EffectsDefinition.ShadowDefinition.BlurRadius = Dimension.Parse(blurStr);

                            string? distanceStr = shadowNode.Attribute("distance")?.Value;
                            if (!string.IsNullOrEmpty(distanceStr))
                                chartDefinition.Grid.HorizontalFormat.EffectsDefinition.ShadowDefinition.Distance = Dimension.Parse(distanceStr);

                            ColorA shadowColor = ColorA.Parse(shadowNode, "color", "transparency");
                            if (shadowColor != null)
                                chartDefinition.Grid.HorizontalFormat.EffectsDefinition.ShadowDefinition.Color = shadowColor;
                        }

                        XElement? glowNode = effectsNode.Element("glow");
                        if (glowNode != null)
                        {
                            chartDefinition.Grid.HorizontalFormat.EffectsDefinition.GlowDefinition = new GlowDefinition();

                            ColorA glowColor = ColorA.Parse(glowNode, "color", "transparency");
                            if (glowColor != null)
                                chartDefinition.Grid.HorizontalFormat.EffectsDefinition.GlowDefinition.Color = glowColor;
                        }
                    }
                }

                // Vertical format (benzer yapı)
                XElement? vertFormatNode = gridNode.Element("vertical-format");
                if (vertFormatNode != null)
                {
                    chartDefinition.Grid.VerticalFormat = new FormatDefinition();

                    XElement? lineNode = vertFormatNode.Element("line");
                    if (lineNode != null)
                    {
                        chartDefinition.Grid.VerticalFormat.LineDefinition = new LineDefinition();

                        XElementAttributeGetter.AsBool(lineNode, "visible", out bool visible);
                        chartDefinition.Grid.VerticalFormat.LineDefinition.Visible = visible;

                        Nullable<DashPreset> dashPreset = XElementAttributeGetter.AsEnum<DashPreset>(lineNode, "dash");
                        if (dashPreset.HasValue)
                            chartDefinition.Grid.VerticalFormat.LineDefinition.DashPreset = dashPreset.Value;

                        Nullable<CompoundLinePreset> compoundLine = XElementAttributeGetter.AsEnum<CompoundLinePreset>(lineNode, "compound");
                        if (compoundLine.HasValue)
                            chartDefinition.Grid.VerticalFormat.LineDefinition.CompoundPreset = compoundLine.Value;

                    }

                    XElement? effectsNode = vertFormatNode.Element("effects");
                    if (effectsNode != null)
                    {
                        chartDefinition.Grid.VerticalFormat.EffectsDefinition = new EffectsDefinition();

                        XElement? glowNode = effectsNode.Element("glow");
                        if (glowNode != null)
                        {
                            chartDefinition.Grid.VerticalFormat.EffectsDefinition.GlowDefinition = new GlowDefinition();

                            ColorA glowColor = ColorA.Parse(glowNode, "color", "transparency");
                            if (glowColor != null)
                                chartDefinition.Grid.VerticalFormat.EffectsDefinition.GlowDefinition.Color = glowColor;
                        }
                    }
                }
            }

            //value-labels

            XElement? valueLabelsNode = chartNode.Element("value-labels");
            if (valueLabelsNode != null)
            {
                chartDefinition.ValueLabels = new ValueLabelDefinition();

                Nullable<DataLabelPosition> position = XElementAttributeGetter.AsEnum<DataLabelPosition>(valueLabelsNode, "position");
                if (position.HasValue)
                    chartDefinition.ValueLabels.Position = position.Value;

                if (XElementAttributeGetter.AsBool(valueLabelsNode, "show", out bool show))
                    chartDefinition.ValueLabels.Show = show;

                if (XElementAttributeGetter.AsBool(valueLabelsNode, "show-legend-key", out bool showLegendKey))
                    chartDefinition.ValueLabels.ShowLegendKey = showLegendKey;

                if (XElementAttributeGetter.AsBool(valueLabelsNode, "show-category-name", out bool showCategoryName))
                    chartDefinition.ValueLabels.ShowCategoryName = showCategoryName;

                if (XElementAttributeGetter.AsBool(valueLabelsNode, "show-series-name", out bool showSeriesName))
                    chartDefinition.ValueLabels.ShowSeriesName = showSeriesName;

                if (XElementAttributeGetter.AsBool(valueLabelsNode, "show-percent", out bool showPercent))
                    chartDefinition.ValueLabels.ShowPercent = showPercent;

                if (XElementAttributeGetter.AsBool(valueLabelsNode, "show-bubble-size", out bool showBubbleSize))
                    chartDefinition.ValueLabels.ShowBubbleSize = showBubbleSize;

                XElement? textFormat = valueLabelsNode.Element("text-format");
                if (textFormat != null)
                {
                    chartDefinition.ValueLabels.TextFormat = XElementAttributeGetter.ParseTextFormat(textFormat);
                }

            }

            //layout

            XElement? layoutNode = chartNode.Element("layout");
            if (layoutNode != null)
            {
                chartDefinition.Layout = new ChartLayoutDefinition();

                if (XElementAttributeGetter.AsInt32(layoutNode, "x", out int xValue))
                {
                    int clampedX = clamp(xValue, 0, 100);
                    chartDefinition.Layout.X = clampedX;
                }

                if (XElementAttributeGetter.AsInt32(layoutNode, "y", out int yValue))
                {
                    int clampedY = clamp(yValue, 0, 100);
                    chartDefinition.Layout.Y = clampedY;
                }

                if (XElementAttributeGetter.AsInt32(layoutNode, "width", out int widthValue))
                {
                    int clampedWidth = clamp(widthValue, 0, 100);
                    chartDefinition.Layout.Width = clampedWidth;
                }

                if (XElementAttributeGetter.AsInt32(layoutNode, "height", out int heightValue))
                {
                    int clampedHeight = clamp(heightValue, 0, 100);
                    chartDefinition.Layout.Height = clampedHeight;
                }

                Nullable<LayoutMode> xMode = XElementAttributeGetter.AsEnum<LayoutMode>(layoutNode, "x-mode");
                if (xMode.HasValue)
                    chartDefinition.Layout.XMode = xMode.Value;

                Nullable<LayoutMode> yMode = XElementAttributeGetter.AsEnum<LayoutMode>(layoutNode, "y-mode");
                if (yMode.HasValue)
                    chartDefinition.Layout.YMode = yMode.Value;

                Nullable<LayoutTarget> target = XElementAttributeGetter.AsEnum<LayoutTarget>(layoutNode, "target");
                if (target.HasValue)
                    chartDefinition.Layout.Target = target.Value;
            }

            //plot-area

            XElement? plotAreaNode = chartNode.Element("plot-area");
            if (plotAreaNode != null)
            {
                XElement? format = plotAreaNode.Element("format");
                if (format != null)
                {
                    chartDefinition.PlotAreaFormat = XElementAttributeGetter.ParseFormatNode(format);
                }
            }

            //data-table

            XElement? dataTableNode = chartNode.Element("data-table");
            if (dataTableNode != null)
            {
                chartDefinition.DataTable = new DataTableDefinition();

                if (XElementAttributeGetter.AsBool(dataTableNode, "show", out bool show))
                    chartDefinition.DataTable.Show = show;

                if (XElementAttributeGetter.AsBool(dataTableNode, "show-horizontal-border", out bool showHorizontalBorder))
                    chartDefinition.DataTable.ShowHorizontalBorder = showHorizontalBorder;

                if (XElementAttributeGetter.AsBool(dataTableNode, "show-vertical-border", out bool showVerticalBorder))
                    chartDefinition.DataTable.ShowVerticalBorder = showVerticalBorder;

                if (XElementAttributeGetter.AsBool(dataTableNode, "show-outline-border", out bool showOutlineBorder))
                    chartDefinition.DataTable.ShowOutlineBorder = showOutlineBorder;

                if (XElementAttributeGetter.AsBool(dataTableNode, "show-legend-key", out bool showLegendKey))
                    chartDefinition.DataTable.ShowLegendKey = showLegendKey;

                XElement? formatNode = dataTableNode.Element("format");
                if (formatNode != null)
                {
                    chartDefinition.DataTable.BoxFormat = XElementAttributeGetter.ParseFormatNode(formatNode);
                }

                XElement? textFormatNode = dataTableNode.Element("text-format");
                if (textFormatNode != null)
                {
                    chartDefinition.DataTable.TextFormat = XElementAttributeGetter.ParseTextFormat(textFormatNode);
                }
            }

            //view-3d

            XElement? view3dNode = chartNode.Element("view-3d");
            if (view3dNode != null)
            {
                chartDefinition.ThreeDView = new ThreeDViewDefinition();

                Nullable<ThreeDShape> shape = XElementAttributeGetter.AsEnum<ThreeDShape>(view3dNode, "shape");
                if (shape.HasValue)
                    chartDefinition.ThreeDView.Shape = shape.Value;

                if (XElementAttributeGetter.AsInt32(view3dNode, "rotation-x", out int rotationX))
                    chartDefinition.ThreeDView.RotationX = clamp(rotationX, -90, 90);

                if (XElementAttributeGetter.AsInt32(view3dNode, "rotation-y", out int rotationY))
                    chartDefinition.ThreeDView.RotationY = clamp(rotationY, 0, 360);

                if (XElementAttributeGetter.AsInt32(view3dNode, "perspective", out int perspective))
                    chartDefinition.ThreeDView.Perspective = clamp(perspective, 0, 240);

                if (XElementAttributeGetter.AsInt32(view3dNode, "depth-percent", out int depthPercent))
                    chartDefinition.ThreeDView.DepthPercent = clamp(depthPercent, 20, 2000);

                if (XElementAttributeGetter.AsInt32(view3dNode, "height-percent", out int heightPercent))
                    chartDefinition.ThreeDView.HeightPercent = clamp(heightPercent, 5, 500);

                if (XElementAttributeGetter.AsInt32(view3dNode, "gap-width", out int gapWidth))
                    chartDefinition.ThreeDView.GapWidth = clamp(gapWidth, 0, 500);

                if (XElementAttributeGetter.AsInt32(view3dNode, "gap-depth", out int gapDepth))
                    chartDefinition.ThreeDView.GapDepth = clamp(gapDepth, 0, 500);

                if (XElementAttributeGetter.AsBool(view3dNode, "right-angle-axes", out bool rightAngleAxes))
                    chartDefinition.ThreeDView.RightAngleAxes = rightAngleAxes;

                if (XElementAttributeGetter.AsBool(view3dNode, "show-floor", out bool showFloor))
                    chartDefinition.ThreeDView.ShowFloor = showFloor;

                if (XElementAttributeGetter.AsBool(view3dNode, "show-back-wall", out bool showBackWall))
                    chartDefinition.ThreeDView.ShowBackWall = showBackWall;

                if (XElementAttributeGetter.AsBool(view3dNode, "show-side-wall", out bool showSideWall))
                    chartDefinition.ThreeDView.ShowSideWall = showSideWall;


                XElement? formatNode = view3dNode.Element("format");
                if (formatNode != null)
                {
                    chartDefinition.ThreeDView.DefaultFormat = XElementAttributeGetter.ParseFormatNode(formatNode);
                }

                XElement? floorFormatNode = view3dNode.Element("floor-format");
                if (floorFormatNode != null)
                {
                    chartDefinition.ThreeDView.FloorFormat = XElementAttributeGetter.ParseFormatNode(floorFormatNode);
                }

                XElement? backWallNode = view3dNode.Element("back-wall-format");
                if (backWallNode != null)
                {
                    chartDefinition.ThreeDView.BackWallFormat = XElementAttributeGetter.ParseFormatNode(backWallNode);
                }

                XElement? sideWallFormatNode = view3dNode.Element("side-wall-format");
                if (sideWallFormatNode != null)
                {
                    chartDefinition.ThreeDView.SideWallFormat = XElementAttributeGetter.ParseFormatNode(sideWallFormatNode);
                }

            }

            //series

            foreach (XElement seriesNode in chartNode.Elements("series"))
            {
                SeriesDefinition seriesDefinition = new SeriesDefinition();
                seriesDefinition.Marker = new MarkerDefinition();

                // Series title parsing
                string? seriesTitle = seriesNode.Attribute("title")?.Value ?? "";
                seriesDefinition.Title = new TitleDefinition { Text = seriesTitle };

                Nullable<MarkerType> seriesMarker = XElementAttributeGetter.AsEnum<MarkerType>(seriesNode, "marker-type");
                if (seriesMarker.HasValue)
                    seriesDefinition.Marker.Type = seriesMarker.Value;

                if (XElementAttributeGetter.AsInt32(seriesNode, "marker-size", out int markerSize))
                    seriesDefinition.Marker.Size = clamp(markerSize, 2, 72);

                ColorA color = ColorA.Parse(seriesNode, "color", "transparency");
                if (color != null)
                {
                    seriesDefinition.Format = new FormatDefinition();
                    seriesDefinition.Format.FillDefinition = new FillDefinition();
                    seriesDefinition.Format.FillDefinition.SolidFillDefinition = new SolidDefinition();
                    seriesDefinition.Format.FillDefinition.SolidFillDefinition.Color = color;
                }

                XElement? formatNode = seriesNode.Element("format");
                if (formatNode != null)
                {
                    seriesDefinition.Format = XElementAttributeGetter.ParseFormatNode(formatNode);
                }

                XElement? markerNode = seriesNode.Element("marker");
                if (markerNode != null)
                {
                    XElement? markerFormatNode = markerNode.Element("format");
                    if (markerFormatNode != null)
                    {
                        seriesDefinition.Marker.Format = XElementAttributeGetter.ParseFormatNode(markerFormatNode);
                    }
                }

                XElement? valueLabelNode = seriesNode.Element("value-labels");
                if (valueLabelNode != null)
                {
                    seriesDefinition.ValueLabels = new ValueLabelDefinition();

                    Nullable<DataLabelPosition> dataLabelPosition = XElementAttributeGetter.AsEnum<DataLabelPosition>(valueLabelNode, "position");
                    if (dataLabelPosition.HasValue)
                        seriesDefinition.ValueLabels.Position = dataLabelPosition.Value;

                    if (XElementAttributeGetter.AsBool(valueLabelNode, "show", out bool show))
                        seriesDefinition.ValueLabels.Show = show;

                    if (XElementAttributeGetter.AsBool(valueLabelNode, "show-legend-key", out bool showLegendKey))
                        seriesDefinition.ValueLabels.ShowLegendKey = showLegendKey;

                    if (XElementAttributeGetter.AsBool(valueLabelNode, "show-category-name", out bool showCategoryName))
                        seriesDefinition.ValueLabels.ShowCategoryName = showCategoryName;

                    if (XElementAttributeGetter.AsBool(valueLabelNode, "show-series-name", out bool showSeriesName))
                        seriesDefinition.ValueLabels.ShowSeriesName = showSeriesName;

                    if (XElementAttributeGetter.AsBool(valueLabelNode, "show-percent", out bool showPercent))
                        seriesDefinition.ValueLabels.ShowPercent = showPercent;

                    if (XElementAttributeGetter.AsBool(valueLabelNode, "show-bubble-size", out bool showBubbleSize))
                        seriesDefinition.ValueLabels.ShowBubbleSize = showBubbleSize;

                    XElement? textFormat = valueLabelNode.Element("text-format");
                    if (textFormat != null)
                    {
                        seriesDefinition.ValueLabels.TextFormat = XElementAttributeGetter.ParseTextFormat(textFormat);
                    }

                }

                IEnumerable<XElement> point = seriesNode.Elements("point");
                if (point != null)
                {
                    foreach (XElement pointNode in point)
                    {
                        PointDefinition pointDefinition = new PointDefinition();

                        string? category = XElementAttributeGetter.AsString(pointNode, "category");
                        if (!string.IsNullOrEmpty(category))
                            pointDefinition.Category = category;

                        if (XElementAttributeGetter.AsDouble(pointNode, "value", out double value))
                            pointDefinition.Value = value;

                        if (XElementAttributeGetter.AsDouble(pointNode, "x", out double x))
                            pointDefinition.X = x;

                        if (XElementAttributeGetter.AsDouble(pointNode, "y", out double y))
                            pointDefinition.Y = y;

                        if (XElementAttributeGetter.AsDouble(pointNode, "size", out double size))
                            pointDefinition.Size = size;

                        XElement? pointVolueLabel = pointNode.Element("value-labels");
                        if (pointVolueLabel != null)
                        {
                            pointDefinition.ValueLabels = new ValueLabelDefinition();
                            Nullable<DataLabelPosition> pointDataLabelPosition = XElementAttributeGetter.AsEnum<DataLabelPosition>(pointVolueLabel, "position");
                            if (pointDataLabelPosition.HasValue)
                                pointDefinition.ValueLabels.Position = pointDataLabelPosition.Value;
                            if (XElementAttributeGetter.AsBool(pointVolueLabel, "show", out bool showP))
                                pointDefinition.ValueLabels.Show = showP;
                            if (XElementAttributeGetter.AsBool(pointVolueLabel, "show-legend-key", out bool showLegendKeyP))
                                pointDefinition.ValueLabels.ShowLegendKey = showLegendKeyP;
                            if (XElementAttributeGetter.AsBool(pointVolueLabel, "show-category-name", out bool showCategoryNameP))
                                pointDefinition.ValueLabels.ShowCategoryName = showCategoryNameP;
                            if (XElementAttributeGetter.AsBool(pointVolueLabel, "show-series-name", out bool showSeriesNameP))
                                pointDefinition.ValueLabels.ShowSeriesName = showSeriesNameP;
                            if (XElementAttributeGetter.AsBool(pointVolueLabel, "show-percent", out bool showPercentP))
                                pointDefinition.ValueLabels.ShowPercent = showPercentP;
                            if (XElementAttributeGetter.AsBool(pointVolueLabel, "show-bubble-size", out bool showBubbleSizeP))
                                pointDefinition.ValueLabels.ShowBubbleSize = showBubbleSizeP;

                            XElement? textFormatP = pointVolueLabel.Element("text-format");
                            if (textFormatP != null)
                            {
                                pointDefinition.ValueLabels.TextFormat = XElementAttributeGetter.ParseTextFormat(textFormatP);
                            }
                        }

                        if (pointNode.Element("format") != null)
                        {
                            pointDefinition.PointFormat = XElementAttributeGetter.ParseFormatNode(pointNode.Element("format"));
                        }

                        seriesDefinition.Points.Add(pointDefinition);
                    }
                }

                chartDefinition.Series.Add(seriesDefinition);
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

            //Format Parse
            internal static FormatDefinition ParseFormatNode(XElement formatNode)
            {
                FormatDefinition formatDefinition = new FormatDefinition();

                XElement? fillNode = formatNode.Element("fill");
                if (fillNode != null)
                    formatDefinition.FillDefinition = ParseFill(fillNode);

                XElement? lineNode = formatNode.Element("line");
                if (lineNode != null)
                    formatDefinition.LineDefinition = ParseLine(lineNode);

                XElement? effectsNode = formatNode.Element("effects");
                if (effectsNode != null)
                    formatDefinition.EffectsDefinition = ParseEffects(effectsNode);

                return formatDefinition;
            }
            internal static FillDefinition ParseFill(XElement fillNode)
            {
                FillDefinition fillDefinition = new FillDefinition();

                XElement? solidNode = fillNode.Element("solid");
                if (solidNode != null)
                {
                    fillDefinition.SolidFillDefinition = new SolidDefinition();
                    ColorA color = ColorA.Parse(solidNode, "color", "transparency");
                    if (color != null)
                        fillDefinition.SolidFillDefinition.Color = color;
                }

                XElement? patternNode = fillNode.Element("pattern");
                if (patternNode != null)
                {
                    PatternDefinition patternDefinition = new PatternDefinition();

                    Nullable<PatternFillPreset> patternType = XElementAttributeGetter.AsEnum<PatternFillPreset>(patternNode, "preset");
                    if (patternType.HasValue)
                        patternDefinition.Preset = patternType.Value;

                    ColorA fg = ColorA.Parse(patternNode, "foreground-color", "foreground-transparency");
                    if (fg != null) patternDefinition.ForegroundColor = fg;

                    ColorA bg = ColorA.Parse(patternNode, "background-color", "background-transparency");
                    if (bg != null) patternDefinition.BackgroundColor = bg;

                    fillDefinition.PatternFillDefinition = patternDefinition;
                }

                XElement? gradientNode = fillNode.Element("gradient");
                if (gradientNode != null)
                    fillDefinition.GradientFillDefinition = ParseGradient(gradientNode);

                return fillDefinition;
            }
            internal static LineDefinition ParseLine(XElement lineNode)
            {
                LineDefinition lineDefinition = new LineDefinition();

                XElementAttributeGetter.AsBool(lineNode, "visible", out bool visibleValue);
                lineDefinition.Visible = visibleValue;

                string? widthStr = lineNode.Attribute("width")?.Value;
                if (!string.IsNullOrEmpty(widthStr))
                    lineDefinition.Width = Dimension.Parse(widthStr);

                Nullable<DashPreset> dashPreset = XElementAttributeGetter.AsEnum<DashPreset>(lineNode, "dash");
                if (dashPreset.HasValue) lineDefinition.DashPreset = dashPreset.Value;

                Nullable<CompoundLinePreset> compoundLine = XElementAttributeGetter.AsEnum<CompoundLinePreset>(lineNode, "compound");
                if (compoundLine.HasValue) lineDefinition.CompoundPreset = compoundLine.Value;

                Nullable<LineCapPreset> cap = XElementAttributeGetter.AsEnum<LineCapPreset>(lineNode, "cap");
                if (cap.HasValue) lineDefinition.CapPreset = cap.Value;

                Nullable<LineJoinPreset> join = XElementAttributeGetter.AsEnum<LineJoinPreset>(lineNode, "join");
                if (join.HasValue) lineDefinition.JoinPreset = join.Value;

                if (XElementAttributeGetter.AsInt32(lineNode, "join-miter-limit", out int joinMiterLimit))
                    lineDefinition.JoinMiterLimit = clamp(joinMiterLimit, 1, 500);

                Nullable<LineEndPreset> beginArrowType = XElementAttributeGetter.AsEnum<LineEndPreset>(lineNode, "begin-arrow-type");
                if (beginArrowType.HasValue) lineDefinition.Begin_Arrow_Type = beginArrowType.Value;

                Nullable<LineEndWidthPreset> beginArrowWidth = XElementAttributeGetter.AsEnum<LineEndWidthPreset>(lineNode, "begin-arrow-width");
                if (beginArrowWidth.HasValue) lineDefinition.Begin_Arrow_Width = beginArrowWidth.Value;

                Nullable<LineEndLengthPreset> beginArrowLength = XElementAttributeGetter.AsEnum<LineEndLengthPreset>(lineNode, "begin-arrow-length");
                if (beginArrowLength.HasValue) lineDefinition.Begin_Arrow_Length = beginArrowLength.Value;

                Nullable<LineEndPreset> endArrowType = XElementAttributeGetter.AsEnum<LineEndPreset>(lineNode, "end-arrow-type");
                if (endArrowType.HasValue) lineDefinition.End_Arrow_Type = endArrowType.Value;

                Nullable<LineEndWidthPreset> endArrowWidth = XElementAttributeGetter.AsEnum<LineEndWidthPreset>(lineNode, "end-arrow-width");
                if (endArrowWidth.HasValue) lineDefinition.End_Arrow_Width = endArrowWidth.Value;

                Nullable<LineEndLengthPreset> endArrowLength = XElementAttributeGetter.AsEnum<LineEndLengthPreset>(lineNode, "end-arrow-length");
                if (endArrowLength.HasValue) lineDefinition.End_Arrow_Length = endArrowLength.Value;

                XElement? solidNode = lineNode.Element("solid");
                if (solidNode != null)
                {
                    lineDefinition.SolidLineDefinition = new SolidDefinition();
                    ColorA color = ColorA.Parse(solidNode, "color", "transparency");
                    if (color != null)
                        lineDefinition.SolidLineDefinition.Color = color;
                }

                XElement? gradientNode = lineNode.Element("gradient");
                if (gradientNode != null)
                    lineDefinition.GradientLineDefinition = ParseGradient(gradientNode);

                return lineDefinition;
            }
            internal static GradientDefinition ParseGradient(XElement gradientNode)
            {
                GradientDefinition gradientDefinition = new GradientDefinition();

                if (XElementAttributeGetter.AsInt32(gradientNode, "angle", out int angle))
                    gradientDefinition.Angle = clamp(angle, 0, 360);

                XElementAttributeGetter.AsBool(gradientNode, "scaled", out bool scaled);
                gradientDefinition.Scaled = scaled;

                foreach (XElement stopNode in gradientNode.Elements("stop"))
                {
                    GradientStopDefinition stop = new GradientStopDefinition();

                    if (XElementAttributeGetter.AsInt32(stopNode, "position", out int pos))
                        stop.Position = clamp(pos, 0, 100) * 1000;

                    ColorA? color = ColorA.Parse(stopNode, "color", "transparency");
                    if (color != null)
                        stop.Color = color;

                    if (stop.Color != null)
                        gradientDefinition.Stops.Add(stop);
                }

                return gradientDefinition;
            }
            internal static EffectsDefinition ParseEffects(XElement effectsNode)
            {
                EffectsDefinition effectsDefinition = new EffectsDefinition();

                XElement? shadowNode = effectsNode.Element("shadow");
                if (shadowNode != null)
                    effectsDefinition.ShadowDefinition = ParseShadow(shadowNode);

                XElement? glowNode = effectsNode.Element("glow");
                if (glowNode != null)
                    effectsDefinition.GlowDefinition = ParseGlow(glowNode);

                XElement? softEdgesNode = effectsNode.Element("soft-edges");
                if (softEdgesNode != null)
                {
                    effectsDefinition.SoftEdgesDefinition = new SoftEdgesDefinition();
                    string? sizeStr = softEdgesNode.Attribute("size")?.Value;
                    if (!string.IsNullOrEmpty(sizeStr))
                        effectsDefinition.SoftEdgesDefinition.Size = Dimension.Parse(sizeStr);
                }

                XElement? reflectionNode = effectsNode.Element("reflection");
                if (reflectionNode != null)
                    effectsDefinition.ReflectionDefinition = ParseReflection(reflectionNode);

                XElement? format3dNode = effectsNode.Element("format-3d");
                if (format3dNode != null)
                    effectsDefinition.Format3dDefinition = ParseFormat3d(format3dNode);

                return effectsDefinition;
            }
            internal static ShadowDefinition ParseShadow(XElement shadowNode)
            {
                ShadowDefinition shadowDefinition = new ShadowDefinition();

                Nullable<ShadowType> shadowType = XElementAttributeGetter.AsEnum<ShadowType>(shadowNode, "type");
                if (shadowType.HasValue)
                    shadowDefinition.Type = shadowType.Value;

                Nullable<ShadowPreset> shadowPreset = XElementAttributeGetter.AsEnum<ShadowPreset>(shadowNode, "preset");
                if (shadowPreset.HasValue)
                    shadowDefinition.Preset = shadowPreset.Value;

                string? blurStr = shadowNode.Attribute("blur-radius")?.Value;
                if (!string.IsNullOrEmpty(blurStr))
                    shadowDefinition.BlurRadius = Dimension.Parse(blurStr);

                string? distanceStr = shadowNode.Attribute("distance")?.Value;
                if (!string.IsNullOrEmpty(distanceStr))
                    shadowDefinition.Distance = Dimension.Parse(distanceStr);

                if (XElementAttributeGetter.AsInt32(shadowNode, "angle", out int angle))
                    shadowDefinition.Angle = clamp(angle, 0, 360);

                ColorA shadowColor = ColorA.Parse(shadowNode, "color", "transparency");
                if (shadowColor != null)
                    shadowDefinition.Color = shadowColor;

                return shadowDefinition;
            }
            internal static GlowDefinition ParseGlow(XElement glowNode)
            {
                GlowDefinition glowDefinition = new GlowDefinition();

                ColorA glowColor = ColorA.Parse(glowNode, "color", "transparency");
                if (glowColor != null)
                    glowDefinition.Color = glowColor;

                string? sizeStr = glowNode.Attribute("size")?.Value;
                if (!string.IsNullOrEmpty(sizeStr))
                    glowDefinition.Size = Dimension.Parse(sizeStr);

                return glowDefinition;
            }
            internal static ReflectionDefinition ParseReflection(XElement reflectionNode)
            {
                ReflectionDefinition reflectionDefinition = new ReflectionDefinition();

                if (reflectionNode.Attribute("blur-radius") != null)
                    reflectionDefinition.Blur = Dimension.Parse(reflectionNode.Attribute("blur-radius")!.Value);

                if (reflectionNode.Attribute("distance") != null)
                    reflectionDefinition.Distance = Dimension.Parse(reflectionNode.Attribute("distance")!.Value);

                if (XElementAttributeGetter.AsInt32(reflectionNode, "start-transparency", out int startTransparency))
                    reflectionDefinition.StartTransparency = clamp(startTransparency, 0, 100);

                if (XElementAttributeGetter.AsInt32(reflectionNode, "end-transparency", out int endTransparency))
                    reflectionDefinition.EndTransparency = clamp(endTransparency, 0, 100);

                if (XElementAttributeGetter.AsInt32(reflectionNode, "start-position", out int startPosition))
                    reflectionDefinition.StartPosition = clamp(startPosition, 0, 100);

                if (XElementAttributeGetter.AsInt32(reflectionNode, "end-position", out int endPosition))
                    reflectionDefinition.EndPosition = clamp(endPosition, 0, 100);

                return reflectionDefinition;
            }
            internal static Format3DDefinition ParseFormat3d(XElement format3dNode)
            {
                Format3DDefinition format3DDefinition = new Format3DDefinition();

                Nullable<MaterialPreset> material = XElementAttributeGetter.AsEnum<MaterialPreset>(format3dNode, "material");
                if (material.HasValue)
                    format3DDefinition.Material = material.Value;

                XElement? bevelNode = format3dNode.Element("bevel");
                if (bevelNode != null)
                {
                    format3DDefinition.BevelDefinition = new BevelDefinition();

                    string? topWidthStr = bevelNode.Attribute("top-width")?.Value;
                    if (!string.IsNullOrEmpty(topWidthStr))
                        format3DDefinition.BevelDefinition.TopWidth = Dimension.Parse(topWidthStr);

                    string? topHeightStr = bevelNode.Attribute("top-height")?.Value;
                    if (!string.IsNullOrEmpty(topHeightStr))
                        format3DDefinition.BevelDefinition.TopHeight = Dimension.Parse(topHeightStr);

                    string? bottomWidthStr = bevelNode.Attribute("bottom-width")?.Value;
                    if (!string.IsNullOrEmpty(bottomWidthStr))
                        format3DDefinition.BevelDefinition.BottomWidth = Dimension.Parse(bottomWidthStr);

                    string? bottomHeightStr = bevelNode.Attribute("bottom-height")?.Value;
                    if (!string.IsNullOrEmpty(bottomHeightStr))
                        format3DDefinition.BevelDefinition.BottomHeight = Dimension.Parse(bottomHeightStr);

                    Nullable<BevelPreset> topPreset = XElementAttributeGetter.AsEnum<BevelPreset>(bevelNode, "top-preset");
                    if (topPreset.HasValue)
                        format3DDefinition.BevelDefinition.TopPreset = topPreset.Value;

                    Nullable<BevelPreset> bottomPreset = XElementAttributeGetter.AsEnum<BevelPreset>(bevelNode, "bottom-preset");
                    if (bottomPreset.HasValue)
                        format3DDefinition.BevelDefinition.BottomPreset = bottomPreset.Value;
                }

                XElement? lightingNode = format3dNode.Element("lighting");
                if (lightingNode != null)
                {
                    format3DDefinition.LightingDefinition = new LightingDefinition();

                    Nullable<LightingPreset> lightingPreset = XElementAttributeGetter.AsEnum<LightingPreset>(lightingNode, "preset");
                    if (lightingPreset.HasValue)
                        format3DDefinition.LightingDefinition.Preset = lightingPreset.Value;

                    if (XElementAttributeGetter.AsInt32(lightingNode, "angle", out int lightingAngle))
                        format3DDefinition.LightingDefinition.Angle = clamp(lightingAngle, 0, 360);

                    Nullable<LightingDirection> lightingDirection = XElementAttributeGetter.AsEnum<LightingDirection>(lightingNode, "direction");
                    if (lightingDirection.HasValue)
                        format3DDefinition.LightingDefinition.Direction = lightingDirection.Value;
                }

                return format3DDefinition;
            }

            //Text-Format Parse
            internal static TextFormatDefinition ParseTextFormat(XElement textFormatNode)
            {
                TextFormatDefinition textFormat = new TextFormatDefinition();

                Nullable<HorizontalAlign> horizontalAlign = XElementAttributeGetter.AsEnum<HorizontalAlign>(textFormatNode, "align-horizontal");
                if (horizontalAlign.HasValue)
                    textFormat.AlignHorizontal = horizontalAlign.Value;

                Nullable<VerticalAlign> verticalAlign = XElementAttributeGetter.AsEnum<VerticalAlign>(textFormatNode, "align-vertical");
                if (verticalAlign.HasValue)
                    textFormat.AlignVertical = verticalAlign.Value;

                Nullable<TextWrapPreset> wrapPreset = XElementAttributeGetter.AsEnum<TextWrapPreset>(textFormatNode, "wrap");
                if (wrapPreset.HasValue)
                    textFormat.Wrap = wrapPreset.Value;

                string? marginLeft = textFormatNode.Attribute("margin-left")?.Value;
                if (!string.IsNullOrEmpty(marginLeft))
                    textFormat.MarginLeft = Dimension.Parse(marginLeft);

                string? marginRight = textFormatNode.Attribute("margin-right")?.Value;
                if (!string.IsNullOrEmpty(marginRight))
                    textFormat.MarginRight = Dimension.Parse(marginRight);

                string? marginTop = textFormatNode.Attribute("margin-top")?.Value;
                if (!string.IsNullOrEmpty(marginTop))
                    textFormat.MarginTop = Dimension.Parse(marginTop);

                string? marginBottom = textFormatNode.Attribute("margin-bottom")?.Value;
                if (!string.IsNullOrEmpty(marginBottom))
                    textFormat.MarginBottom = Dimension.Parse(marginBottom);

                if (XElementAttributeGetter.AsInt32(textFormatNode, "rotate", out int rotateValue))
                    textFormat.Rotate = clamp(rotateValue, -90, 90);

                XElement? fillNode = textFormatNode.Element("fill");
                if (fillNode != null)
                    textFormat.TextFillDefinition = ParseFill(fillNode);

                XElement? lineNode = textFormatNode.Element("line");
                if (lineNode != null)
                    textFormat.TextOutlineDefinition = ParseLine(lineNode);

                XElement? effectsNode = textFormatNode.Element("effects");
                if (effectsNode != null)
                    textFormat.EffectsDefinition = ParseEffects(effectsNode);

                XElement? fontNode = textFormatNode.Element("font");
                if (fontNode != null)
                {
                    textFormat.FontDefinition = new FontDefinition();

                    string? family = XElementAttributeGetter.AsString(fontNode, "family");
                    if (family != null)
                        textFormat.FontDefinition.Family = family;

                    string? sizeStr = XElementAttributeGetter.AsString(fontNode, "size");
                    if (!string.IsNullOrEmpty(sizeStr))
                        textFormat.FontDefinition.Size = Dimension.Parse(sizeStr);

                    if (XElementAttributeGetter.AsBool(fontNode, "bold", out bool bold))
                        textFormat.FontDefinition.Bold = bold;

                    if (XElementAttributeGetter.AsBool(fontNode, "italic", out bool italic))
                        textFormat.FontDefinition.Italic = italic;

                    Nullable<TextUnderlineStyle> underlinePreset = XElementAttributeGetter.AsEnum<TextUnderlineStyle>(fontNode, "underline");
                    if (underlinePreset.HasValue)
                        textFormat.FontDefinition.Underline = underlinePreset.Value;

                    Nullable<TextStrikeStyle> strikePreset = XElementAttributeGetter.AsEnum<TextStrikeStyle>(fontNode, "strike");
                    if (strikePreset.HasValue)
                        textFormat.FontDefinition.Strike = strikePreset.Value;

                    Nullable<TextCapsStyle> capsPreset = XElementAttributeGetter.AsEnum<TextCapsStyle>(fontNode, "caps");
                    if (capsPreset.HasValue)
                        textFormat.FontDefinition.Caps = capsPreset.Value;

                    string? kerningStr = fontNode.Attribute("kerning")?.Value;
                    if (!string.IsNullOrEmpty(kerningStr))
                        textFormat.FontDefinition.Kerning = Dimension.Parse(kerningStr);

                    string? spacingStr = fontNode.Attribute("spacing")?.Value;
                    if (!string.IsNullOrEmpty(spacingStr))
                        textFormat.FontDefinition.Spacing = Dimension.Parse(spacingStr);
                }

                return textFormat;
            }


        }
        public static void AddChart(SlidePart slidePart, XElement chartNode, ref uint shapeId)
        {
            ChartDefinition chartDefinition = GetChartDefinitionFromXml(chartNode);

            ChartPart chartPart = slidePart.AddNewPart<ChartPart>();
            C.ChartSpace chartSpace = new C.ChartSpace();
            chartPart.ChartSpace = chartSpace;
            chartPart.ChartSpace.Append(new C.EditingLanguage() { Val = "en-US" });

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

            chartPart.ChartSpace.Save();

            GraphicFrame graphicFrame = slidePart.Slide.CommonSlideData.ShapeTree.AppendChild(new GraphicFrame());
            graphicFrame.NonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties(
                new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "Chart" + Guid.NewGuid() },
                new NonVisualGraphicFrameDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()
            );

            graphicFrame.Transform = new Transform(
                new A.Offset() { X = 1524000L, Y = 1524000L },
                new A.Extents() { Cx = 6096000L, Cy = 4064000L }
            );

            graphicFrame.Graphic = new A.Graphic(
                new A.GraphicData(
                    new C.ChartReference() { Id = slidePart.GetIdOfPart(chartPart) }
                )
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
            );


            slidePart.Slide.Save();
        }
        private static C.Chart getBarChart(ChartDefinition chartDefinition, XElement chartNode, BarDirectionValues direction)
        {

            uint categoryAxisId = getSafeId();
            uint valueAxisId = getSafeId();

            C.PlotArea plotArea = new C.PlotArea();

            AxisPositionValues categoryAxisPosition = direction == BarDirectionValues.Column ? AxisPositionValues.Bottom : AxisPositionValues.Left;
            AxisPositionValues valueAxisPosition = direction == BarDirectionValues.Column ? AxisPositionValues.Left : AxisPositionValues.Bottom;

            if (chartDefinition.ThreeDView is not null)
            {
                Bar3DChart bar3DChart = new Bar3DChart()
                {
                    BarDirection = new C.BarDirection() { Val = direction },
                    BarGrouping = new C.BarGrouping() { Val = mapBarGrouping(chartDefinition.GroupingType ?? GroupingType.Clustered) },
                    VaryColors = new C.VaryColors() { Val = chartDefinition.VaryColors ?? false }
                };

                addAxisIds(bar3DChart, categoryAxisId, valueAxisId);

                applyBar3DOptions(bar3DChart, chartDefinition.ThreeDView);

                //addBarSeries(chartDefinition,  chartNode, bar3DChart);
                addBarSeries(chartDefinition, bar3DChart);

                addDataLabels(bar3DChart, chartDefinition);

                plotArea.Append(bar3DChart);
            }
            else
            {
                BarChart barChart = new BarChart()
                {
                    BarDirection = new BarDirection() { Val = direction },
                    BarGrouping = new BarGrouping() { Val = mapBarGrouping(chartDefinition.GroupingType ?? GroupingType.Clustered) },
                    VaryColors = new VaryColors() { Val = chartDefinition.VaryColors ?? false }
                };

                addAxisIds(barChart, categoryAxisId, valueAxisId);

                //addBarSeries(chartDefinition, chartNode, barChart);
                addBarSeries(chartDefinition, barChart);

                if (chartDefinition.Overlap.HasValue)
                    barChart.Append(new Overlap() { Val = new SByteValue((sbyte)chartDefinition.Overlap.Value) });

                if (chartDefinition.GapWidth.HasValue)
                    barChart.Append(new GapWidth() { Val = new UInt16Value((ushort)chartDefinition.GapWidth.Value) });

                addDataLabels(barChart, chartDefinition);

                plotArea.Append(barChart);
            }

            addCategoryAxis(chartDefinition, chartNode, plotArea, categoryAxisId, valueAxisId, categoryAxisPosition);
            addValueAxis(chartDefinition, chartNode, plotArea, categoryAxisId, valueAxisId, valueAxisPosition);
            addLayout(chartDefinition.Layout, plotArea);
            //addWallsAndFloor(plotArea, chartDefinition.ThreeDViewDefinition);

            return GetChartCommon(chartDefinition, chartNode, plotArea);
        }
        private static C.Chart getLineChart(ChartDefinition chartDefinition, XElement chartNode)
        {
            uint catId = getSafeId();
            uint valId = getSafeId();

            C.PlotArea plotArea = new();

            bool use3D = chartDefinition.ThreeDView is not null;

            if (use3D)
            {
                Line3DChart line3DChart = new Line3DChart()
                {
                    Grouping = new Grouping() { Val = mapGrouping(chartDefinition.GroupingType ?? GroupingType.Standard) },
                    VaryColors = new VaryColors() { Val = chartDefinition.VaryColors ?? false }
                };

                addAxisIds(line3DChart, catId, valId);
                //addLineSeries(chartDefinition, chartNode, line3DChart);
                addLineSeries(chartDefinition, line3DChart);
                addDataLabels(line3DChart, chartDefinition);

                plotArea.Append(line3DChart);

                addCategoryAxis(chartDefinition, chartNode, plotArea, catId, valId);
                addValueAxis(chartDefinition, chartNode, plotArea, catId, valId);
                addLayout(chartDefinition.Layout, plotArea);

                return GetChartCommon(chartDefinition, chartNode, plotArea);
            }
            else
            {
                LineChart lineChart = new LineChart()
                {
                    Grouping = new Grouping() { Val = mapGrouping(chartDefinition.GroupingType ?? GroupingType.Standard) },
                    VaryColors = new VaryColors() { Val = chartDefinition.VaryColors ?? false }
                };

                addAxisIds(lineChart, catId, valId);

                //addLineSeries(chartDefinition, chartNode, lineChart);
                addLineSeries(chartDefinition, lineChart);
                addDataLabels(lineChart, chartDefinition);

                plotArea.Append(lineChart);

                addCategoryAxis(chartDefinition, chartNode, plotArea, catId, valId);
                addValueAxis(chartDefinition, chartNode, plotArea, catId, valId);
                addLayout(chartDefinition.Layout, plotArea);

                return GetChartCommon(chartDefinition, chartNode, plotArea);
            }
        }
        private static C.Chart getAreaChart(ChartDefinition chartDefinition, XElement chartNode)
        {
            uint catId = getSafeId();
            uint valId = getSafeId();

            C.PlotArea plotArea = new();

            bool use3D = chartDefinition.ThreeDView is not null;

            if (use3D)
            {
                Area3DChart area3DChart = new()
                {
                    Grouping = new Grouping() { Val = mapGrouping(chartDefinition.GroupingType ?? GroupingType.Standard) },
                    VaryColors = new VaryColors() { Val = chartDefinition.VaryColors ?? false }
                };

                addAxisIds(area3DChart, catId, valId);
                //addAreaSeries(chartDefinition, chartNode, area3DChart);
                addAreaSeries(chartDefinition, area3DChart);
                addDataLabels(area3DChart, chartDefinition);

                plotArea.Append(area3DChart);

                addCategoryAxis(chartDefinition, chartNode, plotArea, catId, valId);
                addValueAxis(chartDefinition, chartNode, plotArea, catId, valId);
                addLayout(chartDefinition.Layout, plotArea);

                return GetChartCommon(chartDefinition, chartNode, plotArea);
            }
            else
            {
                AreaChart areaChart = new()
                {
                    Grouping = new Grouping() { Val = mapGrouping(chartDefinition.GroupingType ?? GroupingType.Standard) },
                    VaryColors = new VaryColors() { Val = chartDefinition.VaryColors ?? false }
                };

                addAxisIds(areaChart, catId, valId);
                //addAreaSeries(chartDefinition, chartNode, areaChart);
                addAreaSeries(chartDefinition, areaChart);
                addDataLabels(areaChart, chartDefinition);

                plotArea.Append(areaChart);

                addCategoryAxis(chartDefinition, chartNode, plotArea, catId, valId);
                addValueAxis(chartDefinition, chartNode, plotArea, catId, valId);
                addLayout(chartDefinition.Layout, plotArea);

                return GetChartCommon(chartDefinition, chartNode, plotArea);
            }
        }
        private static C.Chart getPieChart(ChartDefinition chartDefinition, XElement chartNode)
        {

            C.PlotArea plotArea = new C.PlotArea();

            bool use3D = chartDefinition.ThreeDView is not null;

            if (use3D)
            {
                C.Pie3DChart pie3D = new(new VaryColors() { Val = chartDefinition.VaryColors ?? true });

                //addPieSeries(chartDefinition, chartNode, pie3D);
                addPieSeries(chartDefinition, pie3D);
                addDataLabels(pie3D, chartDefinition);

                if (chartDefinition.FirstSliceAngle.HasValue)
                    pie3D.Append(new C.FirstSliceAngle() { Val = (UInt16Value)(ushort)clamp(chartDefinition.FirstSliceAngle.Value, 0, 360) });

                plotArea.Append(pie3D);

                addLayout(chartDefinition.Layout, plotArea);

                // Pie3D'de view3D istemiyoruz
                return GetChartCommon(chartDefinition, chartNode, plotArea, allowView3D: false);
            }

            else
            {
                C.PieChart pieChart = new(new VaryColors() { Val = chartDefinition.VaryColors ?? true });

                //addPieSeries(chartDefinition, chartNode, pieChart);
                addPieSeries(chartDefinition, pieChart);
                addDataLabels(pieChart, chartDefinition);

                if (chartDefinition.FirstSliceAngle.HasValue)
                    pieChart.Append(new C.FirstSliceAngle() { Val = (UInt16Value)(ushort)clamp(chartDefinition.FirstSliceAngle.Value, 0, 360) });

                plotArea.Append(pieChart);

                addLayout(chartDefinition.Layout, plotArea);

                return GetChartCommon(chartDefinition, chartNode, plotArea);
            }

        }
        private static C.Chart getDoughnutChart(ChartDefinition chartDefinition, XElement chartNode)
        {
            C.PlotArea plotArea = new();

            C.DoughnutChart doughnut = new(
                new VaryColors() { Val = chartDefinition.VaryColors ?? true }
            );

            //addPieSeries(chartDefinition, chartNode, doughnut);
            addPieSeries(chartDefinition, doughnut);
            addDataLabels(doughnut, chartDefinition);

            if (chartDefinition.FirstSliceAngle.HasValue)
                doughnut.Append(new C.FirstSliceAngle() { Val = (UInt16Value)(ushort)clamp(chartDefinition.FirstSliceAngle.Value, 0, 360) });

            doughnut.Append(new C.HoleSize() { Val = (ByteValue)(byte)clamp(chartDefinition.DoughnutHoleSize ?? 50, 10, 90) });

            plotArea.Append(doughnut);

            addLayout(chartDefinition.Layout, plotArea);

            return GetChartCommon(chartDefinition, chartNode, plotArea);
        }
        private static C.Chart getScatterChart(ChartDefinition chartDefinition, XElement chartNode)
        {
            uint xId = getSafeId();
            uint yId = getSafeId();

            C.PlotArea plotArea = new();
            C.ScatterStyleValues scatterStyle = mapScatterStyle(chartDefinition.ScatterStyle);

            C.ScatterChart scatter = new(
                new C.ScatterStyle() { Val = scatterStyle },
                new C.VaryColors() { Val = chartDefinition.VaryColors ?? false }
            );

            //addScatterSeries(chartDefinition, chartNode, scatter, scatterStyle);
            addScatterSeries(chartDefinition, scatter, scatterStyle);
            addDataLabels(scatter, chartDefinition);
            addAxisIds(scatter, xId, yId);

            plotArea.Append(scatter);

            // X = Bottom, Y = Left (numerik çift eksen)
            addValueAxis(chartDefinition, chartNode, plotArea, yId, xId, position: AxisPositionValues.Bottom); // X
            addValueAxis(chartDefinition, chartNode, plotArea, xId, yId, position: AxisPositionValues.Left); // Y

            addLayout(chartDefinition.Layout, plotArea);

            return GetChartCommon(chartDefinition, chartNode, plotArea);

        }
        private static C.Chart getBubbleChart(ChartDefinition chartDefinition, XElement chartNode)
        {
            // İki value axis ID’si (X ve Y)
            uint xId = getSafeId();
            uint yId = getSafeId();

            C.PlotArea plotArea = new();

            C.BubbleChart bubbleChart = new(
                new VaryColors() { Val = chartDefinition.VaryColors ?? false }
            );

            //addBubbleSeries(chartDefinition, chartNode, bubbleChart);
            addBubbleSeries(chartDefinition, bubbleChart);
            addDataLabels(bubbleChart, chartDefinition);

            if (chartDefinition.BubbleScale.HasValue)
                bubbleChart.Append(new C.BubbleScale() { Val = (UInt32Value)(uint)clamp(chartDefinition.BubbleScale.Value, 0, 300) });

            bubbleChart.Append(new C.ShowNegativeBubbles() { Val = true });

            addAxisIds(bubbleChart, xId, yId);

            plotArea.Append(bubbleChart);

            addValueAxis(chartDefinition, chartNode, plotArea, yId, xId, position: AxisPositionValues.Bottom);
            addValueAxis(chartDefinition, chartNode, plotArea, xId, yId);

            addLayout(chartDefinition.Layout, plotArea);

            return GetChartCommon(chartDefinition, chartNode, plotArea);
        }
        private static C.Chart getRadarChart(ChartDefinition chartDefinition, XElement chartNode)
        {
            uint catId = getSafeId();
            uint valId = getSafeId();

            C.PlotArea plotArea = new();

            C.RadarChart radar = new(
                new C.RadarStyle { Val = mapRadarStyle(chartDefinition.RadarStyle) },
                new C.VaryColors { Val = chartDefinition.VaryColors ?? false }
            );

            //addRadarSeries(chartDefinition, chartNode, radar);
            addRadarSeries(chartDefinition, radar);

            addAxisIds(radar, catId, valId);
            plotArea.Append(radar);

            addCategoryAxis(chartDefinition, chartNode, plotArea, catId, valId, AxisPositionValues.Bottom);
            addValueAxis(chartDefinition, chartNode, plotArea, catId, valId, AxisPositionValues.Left);

            addLayout(chartDefinition.Layout, plotArea);

            return GetChartCommon(chartDefinition, chartNode, plotArea);
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
            chart.Append(new DisplayBlanksAs() { Val = mapBlanksDisplayedAs(chartDefinition.BlanksDisplayedAs ?? BlanksDisplayedAs.Gap) });
            chart.Append(new AutoTitleDeleted() { Val = !chartDefinition.AutoTitle ?? false });
            chart.Append(new PlotVisibleOnly() { Val = chartDefinition.PlotVisibleOnly ?? false });
            chart.Append(new ShowDataLabelsOverMaximum() { Val = chartDefinition.ShowDataLabelsOverMaximum ?? true });

            if (chartDefinition.Legend is LegendDefinition legendDefinition && legendDefinition.Show)
            {
                C.Legend legend = new C.Legend(
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

                        C.LegendEntry legendEntry = new(
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
                C.DataTable dataTable = new(
                    new C.ShowHorizontalBorder() { Val = dataTableDefinition.ShowHorizontalBorder },
                    new C.ShowVerticalBorder() { Val = dataTableDefinition.ShowVerticalBorder },
                    new C.ShowOutlineBorder() { Val = dataTableDefinition.ShowOutlineBorder },
                    new C.ShowKeys() { Val = dataTableDefinition.ShowLegendKey }
                    );

                if (dataTableDefinition.BoxFormat is not null)
                    applyFormat(dataTable, dataTableDefinition.BoxFormat);   // c:spPr

                if (dataTableDefinition.TextFormat is not null)
                {
                    C.TextProperties txPr = new();
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
        private static void applyBar3DOptions(C.Bar3DChart bar3DChart, ThreeDViewDefinition viewDefinition)
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
        private static void addBarSeries(ChartDefinition chartDefinition, OpenXmlCompositeElement barChart)
        {
            uint seriesIndex = 0;

            foreach (SeriesDefinition seriesDefinition in chartDefinition.Series)
            {
                BarChartSeries series = new BarChartSeries(
                    new C.Index() { Val = seriesIndex },
                    new Order() { Val = seriesIndex },
                    new SeriesText(new C.NumericValue() { Text = seriesDefinition.Title })
                );

                CategoryAxisData catAxisData = new CategoryAxisData();
                C.Values values = new C.Values();

                StringLiteral stringLiteral = new StringLiteral();
                NumberLiteral numberLiteral = new NumberLiteral();

                stringLiteral.Append(new PointCount() { Val = (uint)seriesDefinition.Points.Count });

                for (int i = 0; i < seriesDefinition.Points.Count; i++)
                {
                    stringLiteral.Append(new StringPoint() { Index = (uint)i, NumericValue = new C.NumericValue(seriesDefinition.Points[i].Category) });
                    numberLiteral.Append(new NumericPoint() { Index = (uint)i, NumericValue = new C.NumericValue(seriesDefinition.Points[i].Value?.ToString(CultureInfo.InvariantCulture)) });
                }

                catAxisData.Append(stringLiteral);
                values.Append(numberLiteral);

                applyPointLevelStyling(seriesDefinition, series, useMarker: false);

                series.Append(catAxisData);
                series.Append(values);

                applySeriesFormat(chartDefinition, seriesDefinition, series);
                addDataLabels(series, seriesDefinition);

                barChart.Append(series);

                seriesIndex++;
            }
        }
        private static void addLineSeries(ChartDefinition chartDefinition, OpenXmlCompositeElement lineChart)
        {
            uint idx = 0;

            foreach (SeriesDefinition seriesDefinition in chartDefinition.Series)
            {
                LineChartSeries series = new(
                    new C.Index() { Val = idx },
                    new Order() { Val = idx },
                    new SeriesText(new C.NumericValue() { Text = seriesDefinition.Title })
                );

                buildCategoryAndValues(seriesDefinition, out CategoryAxisData cat, out C.Values vals);

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
        private static void addAreaSeries(ChartDefinition chartDefinition, OpenXmlCompositeElement areaChart)
        {
            uint idx = 0;

            foreach (SeriesDefinition seriesDefinition in chartDefinition.Series)
            {
                AreaChartSeries series = new(
                    new C.Index() { Val = idx },
                    new Order() { Val = idx },
                    new SeriesText(new C.NumericValue() { Text = seriesDefinition.Title })
                );

                buildCategoryAndValues(seriesDefinition, out CategoryAxisData cat, out C.Values vals);

                series.Append(cat);
                series.Append(vals);

                applySeriesFormat(chartDefinition, seriesDefinition, series);
                addDataLabels(series, seriesDefinition);

                areaChart.Append(series);

                idx++;
            }
        }
        private static void addPieSeries(ChartDefinition chartDefinition, OpenXmlCompositeElement owner)
        {
            uint idx = 0;

            foreach (SeriesDefinition seriesDefinition in chartDefinition.Series)
            {
                C.PieChartSeries series = new(
                    new C.Index() { Val = idx },
                    new Order() { Val = idx },
                    new SeriesText(new C.NumericValue() { Text = seriesDefinition.Title })
                );

                buildCategoryAndValues(seriesDefinition, out CategoryAxisData cat, out C.Values vals);
                applyPointLevelStyling(seriesDefinition, series, useMarker: false);

                series.Append(cat);
                series.Append(vals);

                applySeriesFormat(chartDefinition, seriesDefinition, series);
                addDataLabels(series, seriesDefinition);

                owner.Append(series);

                idx++;
            }
        }
        private static void addBubbleSeries(ChartDefinition chartDefinition, C.BubbleChart bubbleChart)
        {
            uint idx = 0;

            foreach (SeriesDefinition seriesDefinition in chartDefinition.Series)
            {
                C.BubbleChartSeries series = new(
                    new C.Index() { Val = idx },
                    new Order() { Val = idx },
                    new SeriesText(new C.NumericValue() { Text = seriesDefinition.Title })
                );

                bool bubble3DEnabled = chartDefinition.Bubble3D ?? false;

                if (bubble3DEnabled)
                    series.Append(new C.InvertIfNegative() { Val = true });

                NumberLiteral xNumLit = new();
                xNumLit.Append(new PointCount() { Val = (uint)seriesDefinition.Points.Count });
                for (int i = 0; i < seriesDefinition.Points.Count; i++)
                    xNumLit.Append(new NumericPoint() { Index = (uint)i, NumericValue = new C.NumericValue(seriesDefinition.Points[i].X?.ToString(CultureInfo.InvariantCulture)) });

                NumberLiteral yNumLit = new();
                yNumLit.Append(new PointCount() { Val = (uint)seriesDefinition.Points.Count });
                for (int i = 0; i < seriesDefinition.Points.Count; i++)
                    yNumLit.Append(new NumericPoint() { Index = (uint)i, NumericValue = new C.NumericValue(seriesDefinition.Points[i].Y?.ToString(CultureInfo.InvariantCulture)) });

                NumberLiteral szNumLit = new();
                szNumLit.Append(new PointCount() { Val = (uint)seriesDefinition.Points.Count });
                for (int i = 0; i < seriesDefinition.Points.Count; i++)
                    szNumLit.Append(new NumericPoint() { Index = (uint)i, NumericValue = new C.NumericValue(seriesDefinition.Points[i].Size?.ToString(CultureInfo.InvariantCulture)) });

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
        private static void addScatterSeries(ChartDefinition chartDefinition, C.ScatterChart scatterChart, C.ScatterStyleValues scatterStyle)
        {
            bool wantsLine =
                scatterStyle == C.ScatterStyleValues.Line ||
                scatterStyle == C.ScatterStyleValues.LineMarker ||
                scatterStyle == C.ScatterStyleValues.Smooth ||
                scatterStyle == C.ScatterStyleValues.SmoothMarker;

            bool wantsMarker =
                scatterStyle == C.ScatterStyleValues.Marker ||
                scatterStyle == C.ScatterStyleValues.LineMarker ||
                scatterStyle == C.ScatterStyleValues.SmoothMarker;

            bool smooth =
                scatterStyle == C.ScatterStyleValues.Smooth ||
                scatterStyle == C.ScatterStyleValues.SmoothMarker;

            uint idx = 0;

            foreach (SeriesDefinition seriesDefinition in chartDefinition.Series)
            {
                C.ScatterChartSeries series = new(
                    new C.Index() { Val = idx },
                    new C.Order() { Val = idx },
                    new C.SeriesText(new C.NumericValue() { Text = seriesDefinition.Title })
                );

                // X
                C.NumberLiteral xNum = new();
                xNum.Append(new C.PointCount() { Val = (uint)seriesDefinition.Points.Count });
                for (int i = 0; i < seriesDefinition.Points.Count; i++)
                    xNum.Append(new C.NumericPoint() { Index = (uint)i, NumericValue = new C.NumericValue(seriesDefinition.Points[i].X?.ToString(CultureInfo.InvariantCulture)) });

                // Y
                C.NumberLiteral yNum = new();
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
        private static void addRadarSeries(ChartDefinition chartDefinition, C.RadarChart radarChart)
        {
            uint idx = 0;

            foreach (SeriesDefinition seriesDefinition in chartDefinition.Series)
            {
                C.RadarChartSeries series = new(
                    new C.Index { Val = idx },
                    new C.Order { Val = idx },
                    new C.SeriesText(new C.NumericValue { Text = seriesDefinition.Title })
                );

                buildCategoryAndValues(seriesDefinition, out var cat, out var vals);

                switch (chartDefinition.RadarStyle)
                {
                    case RadarStyle.Marker:
                        // çizgi + marker
                        ensureShapeLineVisible(series);
                        ensureMarker(series, chartDefinition, seriesDefinition);
                        applyPointLevelStyling(seriesDefinition, series, useMarker: true);
                        applySeriesFormat(chartDefinition, seriesDefinition, series, toOutline: true); // outline
                        break;
                    case RadarStyle.Filled:
                        // alan dolu, marker istemiyoruz, çizgi opsiyonel (outline)
                        ensureNoMarker(series);
                        ensureShapeLineVisible(series);

                        // Seri rengini hem doldurma hem outline için uygula
                        // (toOutline=false -> fill; true -> line)
                        applySeriesFormat(chartDefinition, seriesDefinition, series); // fill
                        applySeriesFormat(chartDefinition, seriesDefinition, series, toOutline: true); // outline
                        break;
                    case RadarStyle.Standard:
                    default:
                        // sadece çizgi, marker yok
                        ensureNoMarker(series);
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
            if (chartDefinition is null || plotArea is null)
                return;

            AxisPositionValues axisPos = position ?? AxisPositionValues.Left;

            ValueAxis valAx = new(
                new C.AxisId() { Val = valueAxisId },
                new Scaling(new Orientation() { Val = C.OrientationValues.MinMax }),
                new AxisPosition() { Val = axisPos },
                new C.MajorTickMark() { Val = chartDefinition.Type == ChartType.Radar ? C.TickMarkValues.Cross : C.TickMarkValues.Outside },
                new C.MinorTickMark() { Val = C.TickMarkValues.None },
                new C.NumberingFormat() { FormatCode = "General", SourceLinked = true },
                new TickLabelPosition() { Val = TickLabelPositionValues.NextTo },
                new CrossingAxis() { Val = categoryAxisId },
                new Crosses() { Val = CrossesValues.AutoZero },
                new CrossBetween() { Val = CrossBetweenValues.Between }
            );

            AxisDefinition targetAxisDefinition = chartDefinition.ValueAxis;
            bool showAxis = chartDefinition.ShowValueAxis;
            bool addMajorGridLines;
            FormatDefinition targetMajorGridlinesFormat = null;

            switch (chartDefinition.Type)
            {
                case ChartType.Bar:
                    addMajorGridLines = chartDefinition.Grid?.ShowVertical == true;
                    targetMajorGridlinesFormat = chartDefinition.Grid?.VerticalFormat;
                    break;
                case ChartType.Column:
                case ChartType.Line:
                case ChartType.Area:
                case ChartType.Radar:
                    addMajorGridLines = chartDefinition.Grid?.ShowHorizontal == true;
                    targetMajorGridlinesFormat = chartDefinition.Grid?.HorizontalFormat;
                    break;
                case ChartType.Bubble:
                case ChartType.Scatter:
                    if (position.HasValue && (position.Value == AxisPositionValues.Left || position.Value == AxisPositionValues.Right))
                    {
                        addMajorGridLines = chartDefinition.Grid?.ShowHorizontal == true;
                        targetMajorGridlinesFormat = chartDefinition.Grid?.HorizontalFormat;
                    }
                    else
                    {
                        addMajorGridLines = chartDefinition.Grid?.ShowVertical == true;
                        targetAxisDefinition = chartDefinition.CategoryAxis;
                        showAxis = chartDefinition.ShowCategoryAxis;
                        targetMajorGridlinesFormat = chartDefinition.Grid?.VerticalFormat;
                    }
                    break;
                case ChartType.Pie:
                case ChartType.Doughnut:
                default:
                    addMajorGridLines = false;
                    break;
            }

            applyAxisDefinition(valAx, targetAxisDefinition);

            if (addMajorGridLines)
                addMajorGridlines(valAx, targetMajorGridlinesFormat);

            valAx.InsertAt(new Delete() { Val = !showAxis }, 2);

            plotArea.Append(valAx);
        }
        private static void addMajorGridlines(OpenXmlCompositeElement owner, FormatDefinition gridFormatDefinition)
        {
            if (owner is null)
                return;

            MajorGridlines majorGridlines = new();

            if (gridFormatDefinition is not null)
                applyFormat(majorGridlines, gridFormatDefinition);

            owner.Append(majorGridlines);
        }
        private static void addCategoryAxis(ChartDefinition chartDefinition, XElement chartNode, C.PlotArea plotArea, uint categoryAxisId, uint valueAxisId, AxisPositionValues? position = null)
        {
            if (chartDefinition is null || plotArea is null)
                return;

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

            bool addMajorGridLines;
            FormatDefinition targetMajorGridlinesFormat = null;

            switch (chartDefinition.Type)
            {
                case ChartType.Bar:
                    addMajorGridLines = chartDefinition.Grid?.ShowHorizontal == true;
                    targetMajorGridlinesFormat = chartDefinition.Grid?.HorizontalFormat;
                    break;
                case ChartType.Column:
                case ChartType.Line:
                case ChartType.Area:
                case ChartType.Radar:
                    addMajorGridLines = chartDefinition.Grid?.ShowVertical == true;
                    targetMajorGridlinesFormat = chartDefinition.Grid?.VerticalFormat;
                    break;
                case ChartType.Pie:
                case ChartType.Doughnut:
                case ChartType.Bubble:
                case ChartType.Scatter:
                default:
                    addMajorGridLines = false;
                    break;
            }

            if (addMajorGridLines)
                addMajorGridlines(catAxis, targetMajorGridlinesFormat);

            plotArea.Append(catAxis);
        }
        private static void addDataLabels(OpenXmlCompositeElement owner, IValueLabelsContainer valueLabelsContainer)
        {
            if (owner is null || valueLabelsContainer is null)
                return;

            ChartDefinition chartDefinition = valueLabelsContainer.GetChartDefinition();

            if (chartDefinition is null)
                return;

            ValueLabelDefinition valueLabelsDefinition = valueLabelsContainer.ValueLabels;

            if (valueLabelsDefinition is null || !valueLabelsDefinition.Show)
                return;

            C.DataLabels dataLabels = new();

            dataLabels.Append(new C.ShowValue() { Val = true });

            DataLabelPositionValues? position = mapDataLabelPosition(valueLabelsDefinition.Position, chartDefinition);

            if (position is not null)
            {
                C.DataLabelPosition dataLabelPosition = new C.DataLabelPosition() { Val = position.Value };
                dataLabels.Append(dataLabelPosition);
            }

            bool showPercent = false;
            bool showBubbleSize = false;
            bool showLeaderLines = false;

            switch (chartDefinition.Type)
            {
                case ChartType.Pie:
                    showPercent = valueLabelsDefinition.ShowPercent;
                    showLeaderLines = position == DataLabelPositionValues.OutsideEnd || position == DataLabelPositionValues.BestFit;
                    break;
                case ChartType.Doughnut:
                    showPercent = valueLabelsDefinition.ShowPercent;
                    showLeaderLines = true;
                    break;
                case ChartType.Scatter:
                    break;
                case ChartType.Bubble:
                    showBubbleSize = valueLabelsDefinition.ShowBubbleSize;
                    break;
                case ChartType.Radar:
                    break;
                default:
                    break;
            }

            dataLabels.Append(new C.ShowLegendKey() { Val = valueLabelsDefinition.ShowLegendKey });
            dataLabels.Append(new C.ShowCategoryName() { Val = valueLabelsDefinition.ShowCategoryName });
            dataLabels.Append(new C.ShowSeriesName() { Val = valueLabelsDefinition.ShowSeriesName });
            dataLabels.Append(new C.ShowPercent() { Val = showPercent });
            dataLabels.Append(new C.ShowBubbleSize() { Val = showBubbleSize });
            dataLabels.Append(new C.ShowLeaderLines() { Val = showLeaderLines });

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
                        C.DataLabel pointDataLabel = new(
                            new C.Index() { Val = pointIndex },
                            new C.ShowValue() { Val = true },
                            new C.ShowLegendKey() { Val = pointValueLabelDefinition.ShowLegendKey },
                            new C.ShowCategoryName() { Val = pointValueLabelDefinition.ShowCategoryName },
                            new C.ShowSeriesName() { Val = pointValueLabelDefinition.ShowSeriesName },
                            new C.ShowPercent() { Val = pointValueLabelDefinition.ShowPercent && (chartDefinition.Type == ChartType.Pie || chartDefinition.Type == ChartType.Doughnut) },
                            new C.ShowBubbleSize() { Val = pointValueLabelDefinition.ShowBubbleSize && chartDefinition.Type == ChartType.Bubble }
                        );

                        if (pointValueLabelDefinition.Position is not null)
                        {
                            DataLabelPositionValues? pointLabelPosition = mapDataLabelPosition(pointValueLabelDefinition.Position.Value, chartDefinition);

                            if (pointLabelPosition is not null)
                            {
                                C.DataLabelPosition dataLabelPosition = new() { Val = pointLabelPosition.Value };
                                pointDataLabel.Append(dataLabelPosition);
                            }
                        }

                        if (pointValueLabelDefinition.TextFormat is not null)
                        {
                            C.TextProperties textProperties = new();

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
        private static void addLayout(ChartLayoutDefinition layoutDefinition, C.PlotArea plotArea)
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
                        Val = layoutDefinition.Target.Value == LayoutTarget.Inner
                            ? LayoutTargetValues.Inner
                            : LayoutTargetValues.Outer
                    });

                layout = new(manualLayout);
            }

            plotArea.InsertAt(layout, 0);
            //plotArea.Append(layout);
        }
        private static void applyFormat3D(C.ChartShapeProperties shapeProperties, Format3DDefinition format3d)
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
            A.LightRig lr = new()
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
                        Width = bevelDefinition.TopWidth?.ToEmu(),
                        Height = bevelDefinition.TopHeight?.ToEmu(),
                        Preset = bevelDefinition.TopPreset is null ? null : mapBevel(bevelDefinition.TopPreset.Value)
                    };
                if (bevelDefinition.BottomWidth is not null || bevelDefinition.BottomHeight is not null || bevelDefinition.BottomPreset is not null)
                    sp3d.BevelBottom = new A.BevelBottom
                    {
                        Width = bevelDefinition.BottomWidth?.ToEmu(),
                        Height = bevelDefinition.BottomHeight?.ToEmu(),
                        Preset = bevelDefinition.BottomPreset is null ? null : mapBevel(bevelDefinition.BottomPreset.Value)
                    };
            }
        }
        private static void applySeriesFormat(ChartDefinition chartDefinition, SeriesDefinition seriesDefinition, OpenXmlCompositeElement series, bool toOutline = false)
        {
            if (chartDefinition is null || seriesDefinition is null || series is null || seriesDefinition.Format is null /*|| chartDefinition.VaryColors == true*/)
                return;

            applyFormat(series, seriesDefinition.Format, chartDefinition.SeriesDefaultFormat);
        }
        private static void applyPointLevelStyling(SeriesDefinition seriesDefinition, OpenXmlCompositeElement series, bool useMarker)
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
        private static C.DataPoint buildDataPoint(int index, PointDefinition pointDefinition, bool useMarker)
        {
            C.DataPoint dpt = new(new C.Index() { Val = (uint)index });

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
        private static void buildCategoryAndValues(SeriesDefinition s, out CategoryAxisData catAxisData, out C.Values values)
        {
            catAxisData = new CategoryAxisData();
            values = new C.Values();

            var stringLiteral = new StringLiteral();
            var numberLiteral = new NumberLiteral();

            stringLiteral.Append(new PointCount() { Val = (uint)s.Points.Count });

            for (int i = 0; i < s.Points.Count; i++)
            {
                stringLiteral.Append(new StringPoint()
                {
                    Index = (uint)i,
                    NumericValue = new C.NumericValue(s.Points[i].Category)
                });
                numberLiteral.Append(new NumericPoint()
                {
                    Index = (uint)i,
                    NumericValue = new C.NumericValue(s.Points[i].Value?.ToString(CultureInfo.InvariantCulture))
                });
            }

            catAxisData.Append(stringLiteral);
            values.Append(numberLiteral);
        }
        private static void applyAxisDefinition(OpenXmlCompositeElement axisNode, AxisDefinition axisDefinition)
        {
            if (axisDefinition is null || axisNode is null)
                return;

            Scaling scaling = axisNode.Elements<Scaling>().FirstOrDefault();
            if (scaling is null)
            {
                scaling = new Scaling();
                axisNode.PrependChild(scaling);
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
        private static void add3DView(C.Chart chart, ChartDefinition chartDefinition)
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
                C.View3D view3D = new();

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
                C.PlotArea plotArea = chart.Elements<C.PlotArea>().FirstOrDefault();

                if (plotArea is not null)
                    chart.InsertBefore(view3D, plotArea);
                else
                    chart.Append(view3D);
            }

            switch (chartDefinition.Type)
            {
                case ChartType.Bar:
                case ChartType.Column:
                case ChartType.Line:
                case ChartType.Area:
                    add3DFloorAndWalls(chart, viewDefinition);
                    break;
                default:
                    break;
            }
        }
        private static void add3DFloorAndWalls(C.Chart chart, ThreeDViewDefinition viewDefinition)
        {
            if (chart is null || viewDefinition is null)
                return;

            static T getWallOrFloor<T>(bool show, FormatDefinition formatDefinition, FormatDefinition defaultFormatDefinition) where T : OpenXmlCompositeElement, new()
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
        private static void applyFormat(OpenXmlCompositeElement node, params FormatDefinition[] formats)
        {
            if (node is null || formats is null)
                return;

            C.ChartShapeProperties charShapeProperties = getOrAddChartShapeProps(node);

            if (charShapeProperties is null)
                return;

            FillDefinition getFillDefinition()
            {
                for (int i = 0; i < formats.Length; i++)
                    if (formats[i]?.FillDefinition is FillDefinition fd)
                        return fd;

                return null;
            }
            LineDefinition getLineDefinition()
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
                        outline.Append(new A.HeadEnd
                        {
                            Type = mapLineEndType(lineDefinition.Begin_Arrow_Type.Value),
                            Width = mapLineEndWidth(lineDefinition.Begin_Arrow_Width ?? LineEndWidthPreset.Medium),
                            Length = mapLineEndLength(lineDefinition.Begin_Arrow_Length ?? LineEndLengthPreset.Medium)
                        });

                    if (lineDefinition.End_Arrow_Type is not null)
                        outline.Append(new A.TailEnd
                        {
                            Type = mapLineEndType(lineDefinition.End_Arrow_Type.Value),
                            Width = mapLineEndWidth(lineDefinition.End_Arrow_Width ?? LineEndWidthPreset.Medium),
                            Length = mapLineEndLength(lineDefinition.End_Arrow_Length ?? LineEndLengthPreset.Medium)
                        });
                }
            }

            // EFFECTS (shadow, glow, soft-edges, 3D, reflection)
            if (getEffectsDefintion() is EffectsDefinition effectsDefinition)
                applyEffects(node, effectsDefinition);
        }
        private static void applyEffects(OpenXmlCompositeElement node, EffectsDefinition effectsDefinition)
        {
            if (node is null || effectsDefinition is null)
                return;

            C.ChartShapeProperties chartShapeProperties = getOrAddChartShapeProps(node);

            if (chartShapeProperties is null)
                return;

            A.EffectList effectList = new();
            Dimension placeHolderDimension = new(4, UnitType.Pt);

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
        private static void applyShadow(A.EffectList effectList, ShadowDefinition shadowDefinition, Dimension placeHolderDimension, ShadowType defaultShadowType)
        {
            if (effectList is null || shadowDefinition is null || placeHolderDimension is null)
                return;

            ShadowType shadowType = shadowDefinition.Type ?? defaultShadowType;

            effectList.RemoveAllChildren<A.OuterShadow>();
            effectList.RemoveAllChildren<A.InnerShadow>();
            effectList.RemoveAllChildren<A.PresetShadow>();

            A.RgbColorModelHex shadowHex = toHexFill(shadowDefinition?.Color, "000000");
            Int64Value blurRadius = (shadowDefinition.BlurRadius ?? placeHolderDimension).ToEmu();
            Int64Value distance = (shadowDefinition.Distance ?? placeHolderDimension).ToEmu();
            Int32Value direction = (shadowDefinition.Angle ?? 45) * 60000;
            A.PresetShadowValues preset = mapShadowPreset(shadowDefinition.Preset);

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
        private static void applyGlow(A.EffectList effectList, GlowDefinition glowDefinition, Dimension placeHolderDimension)
        {
            if (effectList is null || glowDefinition is null || placeHolderDimension is null)
                return;

            effectList.RemoveAllChildren<A.Glow>();

            A.Glow glow = new()
            {
                RgbColorModelHex = toHexFill(glowDefinition.Color, "000000"),
                Radius = (glowDefinition.Size ?? placeHolderDimension).ToEmu(),
            };

            effectList.Append(glow);
        }
        private static void applySoftEdges(A.EffectList effectList, SoftEdgesDefinition softEdgesDefinition, Dimension placeHolderDimension)
        {
            if (effectList is null || softEdgesDefinition is null || placeHolderDimension is null)
                return;

            effectList.RemoveAllChildren<A.SoftEdge>();
            effectList.Append(new A.SoftEdge() { Radius = (softEdgesDefinition.Size ?? placeHolderDimension).ToEmu() });
        }
        private static void applyReflection(A.EffectList effectList, ReflectionDefinition reflectionDefinition)
        {
            if (effectList is null || reflectionDefinition is null || !reflectionDefinition.HasData)
                return;

            effectList.RemoveAllChildren<A.Reflection>();

            A.Reflection reflection = new();

            if (reflectionDefinition.Blur is not null) reflection.BlurRadius = reflectionDefinition.Blur.ToEmu();
            if (reflectionDefinition.Distance is not null) reflection.Distance = reflectionDefinition.Distance.ToEmu();
            if (reflectionDefinition.StartTransparency is int st) reflection.StartOpacity = transparencyToAlpha(st);
            if (reflectionDefinition.EndTransparency is int et) reflection.EndAlpha = transparencyToAlpha(et);
            if (reflectionDefinition.StartPosition is int sp) reflection.StartPosition = sp * 1000;
            if (reflectionDefinition.EndPosition is int ep) reflection.EndPosition = ep * 1000;

            effectList.Append(reflection);
        }
        private static void applyFormatJoin(A.Outline outline, LineDefinition lineDefinition)
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
                    outline.Append(new A.Miter() { Limit = (lineDefinition.JoinMiterLimit ?? 800000) });
                    break;
                default:
                    break;
            }
        }
        private static bool applyTextFormat(OpenXmlCompositeElement txPrOwner, TextFormatDefinition textFormatDefinition)
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

            if (textFormatDefinition.MarginLeft is not null)
                bodyPr.LeftInset = textFormatDefinition.MarginLeft.ToEmu();

            if (textFormatDefinition.MarginRight is not null)
                bodyPr.RightInset = textFormatDefinition.MarginRight.ToEmu();

            if (textFormatDefinition.MarginTop is not null)
                bodyPr.TopInset = textFormatDefinition.MarginTop.ToEmu();

            if (textFormatDefinition.MarginBottom is not null)
                bodyPr.BottomInset = textFormatDefinition.MarginBottom.ToEmu();

            if (textFormatDefinition.Rotate is not null)
                bodyPr.Rotation = textFormatDefinition.Rotate.Value * 60000;

            applyTextOutline(defRPr, textFormatDefinition.TextOutlineDefinition);
            applyTextFill(defRPr, textFormatDefinition.TextFillDefinition);
            applyTextEffects(defRPr, textFormatDefinition.EffectsDefinition);
            applyFont(defRPr, textFormatDefinition.FontDefinition);

            if (p.GetFirstChild<A.Run>() is null)
                p.Append(new A.Run(new A.RunProperties(), new A.Text() { Text = string.Empty }));

            return true;
        }
        private static void applyFont(A.TextCharacterPropertiesType runProperties, FontDefinition fontDefinition)
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
        private static void applyTextFill(A.TextCharacterPropertiesType runProperties, FillDefinition fillDefinition)
        {
            if (runProperties is null)
                return;

            runProperties.RemoveAllChildren<A.NoFill>();
            runProperties.RemoveAllChildren<A.SolidFill>();
            runProperties.RemoveAllChildren<A.GradientFill>();
            runProperties.RemoveAllChildren<A.PatternFill>();
            runProperties.RemoveAllChildren<A.BlipFill>();
            runProperties.RemoveAllChildren<A.GroupFill>();

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

            if (fillDefinition.SolidFillDefinition?.Color is not null)
            {
                A.SolidFill sf = new(toHexFill(fillDefinition.SolidFillDefinition.Color, "000000"));

                if (sf is not null)
                    runProperties.Append(sf);
            }
        }
        private static void applyTextOutline(A.TextCharacterPropertiesType runProperties, LineDefinition lineDefinition)
        {
            if (runProperties is null || lineDefinition is null)
                return;

            A.Outline outline = getOrAddOutline(runProperties);

            outline.RemoveAllChildren<A.NoFill>();
            outline.RemoveAllChildren<A.SolidFill>();
            outline.RemoveAllChildren<A.GradientFill>();
            outline.RemoveAllChildren<A.PatternFill>();
            outline.RemoveAllChildren<A.BlipFill>();
            outline.RemoveAllChildren<A.GroupFill>();

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
        private static void applyTextEffects(A.TextCharacterPropertiesType runProperties, EffectsDefinition effectsDefinition)
        {
            if (runProperties is null || effectsDefinition is null)
                return;

            A.EffectList effectList = new();
            Dimension placeHolderDimension = new(1, UnitType.Pt);

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
        private static A.SolidFill buildSolidFill(ColorA colorA, string fallbackColor = null)
        {
            if (colorA is null && string.IsNullOrWhiteSpace(fallbackColor))
                return null;

            A.RgbColorModelHex rgb = toHexFill(colorA, fallback: fallbackColor);

            if (rgb is null)
                return null;

            return new A.SolidFill(rgb);
        }
        private static A.PatternFill buildPatternFill(PatternDefinition patternDefinition)
        {
            if (patternDefinition is null)
                return null;

            A.PatternFill patternFill = new()
            {
                Preset = mapPatternPreset(patternDefinition.Preset),
                ForegroundColor = new A.ForegroundColor(toHexFill(patternDefinition.ForegroundColor, "000000")),
                BackgroundColor = new A.BackgroundColor(toHexFill(patternDefinition.BackgroundColor, "FFFFFF")),
            };

            return patternFill;
        }
        private static A.GradientFill buildGradientFill(GradientDefinition gradientDefinition)
        {
            if (gradientDefinition is null)
                return null;
            bool scaled = gradientDefinition.Scaled ?? true;
            A.GradientFill gradientFill = new() { RotateWithShape = new BooleanValue(!scaled) };
            A.GradientStopList gradientStopList = new();
            List<GradientStopDefinition> gradientStops = gradientDefinition.Stops;

            if (gradientStops == null || gradientStops.Count == 0)
            {
                gradientStops = new List<GradientStopDefinition>();

                gradientStops.Add(new GradientStopDefinition()
                {
                    Position = 0,
                    Color = new ColorA("#000000", null),
                });

                gradientStops.Add(new GradientStopDefinition()
                {
                    Position = 100,
                    Color = new ColorA("#ffffff", null),
                });
            }

            foreach (GradientStopDefinition stopDefinition in gradientStops)
            {
                A.GradientStop gradientStop = new()
                {
                    Position = stopDefinition.Position * 1000, // 0..100000
                    RgbColorModelHex = toHexFill(stopDefinition.Color, "000000"),
                };

                gradientStopList.Append(gradientStop);
            }

            gradientFill.Append(gradientStopList);
            gradientFill.Append(new A.LinearGradientFill()
            {
                Angle = new Int32Value((gradientDefinition.Angle ?? 45) * 60000),
                Scaled = scaled
            });

            return gradientFill;
        }
        private static string ToHexColorCode(string colorOrName)
        {
            if (string.IsNullOrWhiteSpace(colorOrName))
                return null;

            colorOrName = colorOrName.Trim();

            if (colorOrName.StartsWith("#"))
                colorOrName = colorOrName.Substring(1);

            if (System.Text.RegularExpressions.Regex.IsMatch(colorOrName, @"^[0-9a-fA-F]{6,8}$"))
                return colorOrName.ToUpper();

            return null;
        }


        private static A.RgbColorModelHex toHexFill(string hexOrNamed, string fallback = null, int? transparency = null)
        {
            string hexColor = ToHexColorCode(string.IsNullOrWhiteSpace(hexOrNamed) ? fallback : hexOrNamed);

            if (string.IsNullOrWhiteSpace(hexColor))
                return null;

            A.RgbColorModelHex hex = new A.RgbColorModelHex() { Val = hexColor };

            if (transparency is not null)
                hex.Append(new A.Alpha { Val = transparencyToAlpha(transparency.Value) });

            return hex;
        }
        private static A.RgbColorModelHex toHexFill(ColorA colorA, string fallback = null)
        {
            return toHexFill(colorA?.Color, fallback, colorA?.Transparency);
        }
        private static int transparencyToAlpha(int transparencyPercent)
        {
            int t = clamp(transparencyPercent, 0, 100);  // % şeffaflık
            int alphaVal = (int)Math.Round((100 - t) * 1000.0); // %opaklık → 0..100000

            return clamp(alphaVal, 0, 100000);
        }
        private static int clamp(int v, int min, int max)
        {
            if (v < min) return min;
            if (v > max) return max;

            return v;
        }
        private static C.ChartShapeProperties getOrAddChartShapeProps(OpenXmlCompositeElement node)
        {
            return getOrAdd<C.ChartShapeProperties>(node);
        }
        private static C.TextProperties getOrAddTextProperties(OpenXmlCompositeElement node)
        {
            return getOrAdd<C.TextProperties>(node);
        }
        private static A.Outline getOrAddOutline(OpenXmlCompositeElement node)
        {
            return getOrAdd<A.Outline>(node);
        }
        private static C.Marker getOrAddMarker(OpenXmlCompositeElement node)
        {
            return getOrAdd<C.Marker>(node);
        }
        private static T getOrAdd<T>(OpenXmlCompositeElement node) where T : OpenXmlCompositeElement, new()
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
        private static void ensureMarker(OpenXmlCompositeElement series, ChartDefinition chartDefinition, SeriesDefinition seriesDefinition)
        {
            if (series is null || chartDefinition is null || seriesDefinition is null)
                return;

            C.Marker marker = getOrAddMarker(series);

            if (marker is null)
                return;

            MarkerType? markerType = seriesDefinition.Marker?.Type ?? chartDefinition.Marker?.Type;
            int? markerSize = seriesDefinition.Marker?.Size ?? chartDefinition.Marker?.Size;

            if (markerType is not null)
                marker.Symbol = new C.Symbol() { Val = mapMarker(markerType.Value) };

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
                    string seriesColor = seriesSolidColor.Color;
                    bool applyMarkerFillAlpha = chartDefinition.Type == ChartType.Scatter && chartDefinition.ScatterStyle == ScatterStyle.Marker;

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
        private static void ensureNoMarker(OpenXmlCompositeElement series)
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
        private static void ensureShapeLineVisible(OpenXmlCompositeElement node)
        {
            if (node is null)
                return;

            C.ChartShapeProperties spPr = node.GetFirstChild<C.ChartShapeProperties>();

            if (spPr?.GetFirstChild<A.Outline>()?.GetFirstChild<A.NoFill>() is A.NoFill nf)
                nf.Remove();
        }
        private static void ensureShapeLineHidden(OpenXmlCompositeElement node)
        {
            if (node is null)
                return;

            C.ChartShapeProperties spPr = getOrAddChartShapeProps(node);
            A.Outline outline = getOrAddOutline(spPr);

            if (outline.GetFirstChild<A.NoFill>() is null)
                outline.Append(new A.NoFill());
        }
        private static A.GradientFill buildLinearGradientFill(string color1, string color2, int? angleDegrees)
        {
            A.GradientFill gradientFill = new A.GradientFill() { RotateWithShape = new BooleanValue(true) };
            A.LinearGradientFill linearGradientFill = new A.LinearGradientFill() { Angle = new Int32Value((angleDegrees ?? 45) * 60000) };
            A.GradientStopList gradientStopList = new A.GradientStopList();
            A.GradientStop gradientStop1 = new A.GradientStop(toHexFill(color1, "000000")) { Position = new Int32Value(0) };
            A.GradientStop gradientStop2 = new A.GradientStop(toHexFill(color2, "FFFFFF")) { Position = new Int32Value(100000) }; // 0..100000

            gradientStopList.Append(gradientStop1);
            gradientStopList.Append(gradientStop2);
            gradientFill.Append(gradientStopList);
            gradientFill.Append(linearGradientFill);

            return gradientFill;
        }
        private static A.Outline buildBorder(string color, Dimension width)
        {
            A.Outline outline = new A.Outline(new A.SolidFill(toHexFill(color, "000000")));

            if (width is not null)
                outline.Width = width.ToEmu();

            return outline;
        }
        private static A.SolidFill buildSolidFill(string color)
        {
            return new A.SolidFill(toHexFill(color, "dedede"));
        }
        private static A.PatternFill buildPatternFill(PatternFillPreset patternFillPreset, string foreGround, string backGround)
        {
            A.PatternFill p = new A.PatternFill() { Preset = mapPatternPreset(patternFillPreset) };

            p.ForegroundColor = new A.ForegroundColor(toHexFill(foreGround, "000000"));
            p.BackgroundColor = new A.BackgroundColor(toHexFill(backGround, "FFFFFF"));

            return p;
        }
        private static void applyTickLabelRotation(OpenXmlCompositeElement axisNode, int? rotation)
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

        //private static t getwallorfloor<t>(bool show, ıfillandbordercontainer fillandbordercontainer, fillstyle fillstyle, ıfillandbordercontainer defaultfillandbordercontainer = null, int? gradientangleoverride = null) where t : openxmlcompositeelement, new()
        //{
        //    t elem = new t();
        //    if (!show)
        //    {
        //        // görünmez yapmak için: nofill + noline
        //        elem.append(new c.chartshapeproperties(
        //            new a.nofill(),
        //            new a.outline(new a.nofill())
        //        ));
        //    }
        //    else
        //    {
        //        applyfillandborder(elem, fillandbordercontainer, fillstyle, defaultfillandbordercontainer, gradientangleoverride);
        //    }
        //    return elem;
        //}
        //private static void applyfillandborder(openxmlcompositeelement node, ıfillandbordercontainer definition, fillstyle fillstyle, ıfillandbordercontainer defaultdefinition = null, int? gradientangleoverride = null)
        //{
        //    if (node is null || definition is null)
        //        return;

        //    c.chartshapeproperties shapeproperties = new();

        //    bool tryaddborder(formatdefinition container)
        //    {
        //        if (container is null || !container.hasborder())
        //            return false;

        //        shapeproperties.append(buildborder(container.bordercolor, container.borderwidth));

        //        return true;
        //    }
        //    bool tryaddgradientfill(ıfillandbordercontainer container, bool tryfallback)
        //    {
        //        if (container is null)
        //            return false;

        //        if (container.hasgradientfill())
        //        {
        //            shapeproperties.append(buildlineargradientfill(container.gradientfillcolor1, container.gradientfillcolor2, gradientangleoverride ?? container.gradientangle));

        //            return true;
        //        }
        //        else if (tryfallback)
        //        {
        //            return tryaddpatternfill(container, true);
        //        }

        //        return false;
        //    }
        //    bool tryaddpatternfill(ıfillandbordercontainer container, bool tryfallback)
        //    {
        //        if (container is null)
        //            return false;

        //        if (container.haspatternfill())
        //        {
        //            shapeproperties.append(buildpatternfill(container.patternfillpreset.value, container.patternfillforeground, container.patternfillbackground));

        //            return true;
        //        }
        //        else if (tryfallback)
        //        {
        //            return tryaddsolidfill(container);
        //        }

        //        return false;
        //    }
        //    bool tryaddsolidfill(ıfillandbordercontainer container)
        //    {
        //        if (container is null)
        //            return false;

        //        if (container.hassolidfill())
        //        {
        //            shapeproperties.append(buildsolidfill(container.solidfillcolor));

        //            return true;
        //        }

        //        return false;
        //    }

        //    switch (fillstyle)
        //    {
        //        case fillstyle.auto:
        //            if (!tryaddgradientfill(definition, true))
        //                tryaddgradientfill(defaultdefinition, true);
        //            break;
        //        case fillstyle.gradient:
        //            if (!tryaddgradientfill(definition, false))
        //                tryaddgradientfill(defaultdefinition, false);
        //            break;
        //        case fillstyle.pattern:
        //            if (!tryaddpatternfill(definition, false))
        //                tryaddpatternfill(defaultdefinition, false);
        //            break;
        //        case fillstyle.solid:
        //            if (!tryaddsolidfill(definition))
        //                tryaddsolidfill(defaultdefinition);
        //            break;
        //        case fillstyle.none:
        //        default:
        //            break;
        //    }

        //    if (!tryaddborder(definition))
        //        tryaddborder(defaultdefinition);

        //    if (shapeproperties.haschildren)
        //        node.append(shapeproperties);
        //}
        private static C.Floor ensureFloor(OpenXmlCompositeElement plotArea)
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
        private static C.SideWall ensureSideWall(OpenXmlCompositeElement plotArea)
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
        private static C.BackWall ensureBackWall(OpenXmlCompositeElement plotArea)
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
        private static C.DataLabels ensureDataLabels(OpenXmlCompositeElement node)
        {
            C.DataLabels dataLabels = node.GetFirstChild<C.DataLabels>();

            dataLabels ??= node.AppendChild(new C.DataLabels());

            return dataLabels;
        }
        private static GroupingValues mapGrouping(GroupingType value)
        {
            return value switch
            {
                GroupingType.Stacked => GroupingValues.Stacked,
                GroupingType.PercentStacked => GroupingValues.PercentStacked,
                _ => GroupingValues.Standard,
            };
        }
        private static MarkerStyleValues mapMarker(MarkerType value)
        {
            return value switch
            {
                MarkerType.Circle => MarkerStyleValues.Circle,
                MarkerType.Square => MarkerStyleValues.Square,
                MarkerType.Diamond => MarkerStyleValues.Diamond,
                MarkerType.Triangle => MarkerStyleValues.Triangle,
                MarkerType.X => MarkerStyleValues.X,
                _ => MarkerStyleValues.Circle,
            };
        }
        private static C.DisplayBlanksAsValues mapBlanksDisplayedAs(BlanksDisplayedAs value)
        {
            return value switch
            {
                BlanksDisplayedAs.Gap => C.DisplayBlanksAsValues.Gap,
                BlanksDisplayedAs.Zero => C.DisplayBlanksAsValues.Zero,
                BlanksDisplayedAs.Span => C.DisplayBlanksAsValues.Span,
                _ => C.DisplayBlanksAsValues.Gap
            };
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
        private static LegendPositionValues mapLegendPosition(LegendPosition value)
        {
            return value switch
            {
                LegendPosition.Bottom => LegendPositionValues.Bottom,
                LegendPosition.Left => LegendPositionValues.Left,
                LegendPosition.Right => LegendPositionValues.Right,
                LegendPosition.Top => LegendPositionValues.Top,
                LegendPosition.TopRight => LegendPositionValues.TopRight,
                _ => LegendPositionValues.Right,
            };
        }
        private static DataLabelPositionValues? mapDataLabelPosition(DataLabelPosition? value, ChartDefinition chartDefinition)
        {
            if (chartDefinition is null || chartDefinition.ThreeDView is not null || (chartDefinition.Type == ChartType.Bubble && chartDefinition.Bubble3D == true))
                return null;

            ChartType chartType = chartDefinition.Type;

            switch (chartType)
            {
                case ChartType.Bar:
                case ChartType.Column:
                    return value switch
                    {
                        DataLabelPosition.Center => DataLabelPositionValues.Center,
                        DataLabelPosition.InsideEnd => DataLabelPositionValues.InsideEnd,
                        DataLabelPosition.InsideBase => DataLabelPositionValues.InsideBase,
                        DataLabelPosition.OutsideEnd => DataLabelPositionValues.OutsideEnd,
                        _ => DataLabelPositionValues.OutsideEnd,
                    };
                case ChartType.Line:
                case ChartType.Scatter:
                case ChartType.Bubble:
                    return value switch
                    {
                        DataLabelPosition.Center => DataLabelPositionValues.Center,
                        DataLabelPosition.Left => DataLabelPositionValues.Left,
                        DataLabelPosition.Right => DataLabelPositionValues.Right,
                        DataLabelPosition.Top => DataLabelPositionValues.Top,
                        DataLabelPosition.Bottom => DataLabelPositionValues.Bottom,
                        _ => DataLabelPositionValues.Top,
                    };
                case ChartType.Area:
                case ChartType.Radar:
                case ChartType.Doughnut:
                    return null;
                case ChartType.Pie:
                    return value switch
                    {
                        DataLabelPosition.Center => DataLabelPositionValues.Center,
                        DataLabelPosition.InsideEnd => DataLabelPositionValues.InsideEnd,
                        DataLabelPosition.OutsideEnd => DataLabelPositionValues.OutsideEnd,
                        DataLabelPosition.BestFit => DataLabelPositionValues.BestFit,
                        _ => DataLabelPositionValues.BestFit,
                    };
                default:
                    return null;
            }
        }
        private static ShapeValues mapShapeValues(ThreeDShape? value)
        {
            if (value is null)
                return ShapeValues.Box;

            return value.Value switch
            {
                ThreeDShape.ConeToMax => ShapeValues.ConeToMax,
                ThreeDShape.Cone => ShapeValues.Cone,
                ThreeDShape.Cylinder => ShapeValues.Cylinder,
                ThreeDShape.Box => ShapeValues.Box,
                ThreeDShape.Pyramid => ShapeValues.Pyramid,
                ThreeDShape.PyramidToMaximum => ShapeValues.PyramidToMaximum,
                _ => ShapeValues.Box,
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
        private static C.ScatterStyleValues mapScatterStyle(ScatterStyle value)
        {
            return value switch
            {
                ScatterStyle.Line => C.ScatterStyleValues.Line,
                ScatterStyle.LineMarker => C.ScatterStyleValues.LineMarker,
                ScatterStyle.Marker => C.ScatterStyleValues.Marker,
                ScatterStyle.Smooth => C.ScatterStyleValues.Smooth,
                ScatterStyle.SmoothMarker => C.ScatterStyleValues.SmoothMarker,
                _ => C.ScatterStyleValues.Marker,
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
        private static A.BevelPresetValues mapBevel(BevelPreset value)
        {
            return value switch
            {
                BevelPreset.RelaxedInset => A.BevelPresetValues.RelaxedInset,
                BevelPreset.Circle => A.BevelPresetValues.Circle,
                BevelPreset.Slope => A.BevelPresetValues.Slope,
                BevelPreset.Cross => A.BevelPresetValues.Cross,
                BevelPreset.Angle => A.BevelPresetValues.Angle,
                BevelPreset.SoftRound => A.BevelPresetValues.SoftRound,
                BevelPreset.Convex => A.BevelPresetValues.Convex,
                BevelPreset.CoolSlant => A.BevelPresetValues.CoolSlant,
                BevelPreset.Divot => A.BevelPresetValues.Divot,
                BevelPreset.Riblet => A.BevelPresetValues.Riblet,
                BevelPreset.HardEdge => A.BevelPresetValues.HardEdge,
                BevelPreset.ArtDeco => A.BevelPresetValues.ArtDeco,
                _ => A.BevelPresetValues.RelaxedInset
            };
        }
        private static A.PresetMaterialTypeValues mapMaterial(MaterialPreset value)
        {
            return value switch
            {
                MaterialPreset.LegacyMatte => A.PresetMaterialTypeValues.LegacyMatte,
                MaterialPreset.LegacyPlastic => A.PresetMaterialTypeValues.LegacyPlastic,
                MaterialPreset.LegacyMetal => A.PresetMaterialTypeValues.LegacyMetal,
                MaterialPreset.LegacyWireframe => A.PresetMaterialTypeValues.LegacyWireframe,
                MaterialPreset.Matte => A.PresetMaterialTypeValues.Matte,
                MaterialPreset.Plastic => A.PresetMaterialTypeValues.Plastic,
                MaterialPreset.Metal => A.PresetMaterialTypeValues.Metal,
                MaterialPreset.WarmMatte => A.PresetMaterialTypeValues.WarmMatte,
                MaterialPreset.TranslucentPowder => A.PresetMaterialTypeValues.TranslucentPowder,
                MaterialPreset.Powder => A.PresetMaterialTypeValues.Powder,
                MaterialPreset.DarkEdge => A.PresetMaterialTypeValues.DarkEdge,
                MaterialPreset.SoftEdge => A.PresetMaterialTypeValues.SoftEdge,
                MaterialPreset.Clear => A.PresetMaterialTypeValues.Clear,
                MaterialPreset.Flat => A.PresetMaterialTypeValues.Flat,
                MaterialPreset.SoftMetal => A.PresetMaterialTypeValues.SoftMetal,
                _ => A.PresetMaterialTypeValues.Matte
            };
        }
        private static A.LightRigDirectionValues mapLightDir(LightingDirection value)
        {
            return value switch
            {
                LightingDirection.TopLeft => A.LightRigDirectionValues.TopLeft,
                LightingDirection.Top => A.LightRigDirectionValues.Top,
                LightingDirection.TopRight => A.LightRigDirectionValues.TopRight,
                LightingDirection.Left => A.LightRigDirectionValues.Left,
                LightingDirection.Right => A.LightRigDirectionValues.Right,
                LightingDirection.BottomLeft => A.LightRigDirectionValues.BottomLeft,
                LightingDirection.Bottom => A.LightRigDirectionValues.Bottom,
                LightingDirection.BottomRight => A.LightRigDirectionValues.BottomRight,
                _ => A.LightRigDirectionValues.Top
            };
        }
        private static A.LightRigValues mapLightPreset(LightingPreset value)
        {
            return value switch
            {
                LightingPreset.LegacyFlat1 => A.LightRigValues.LegacyFlat1,
                LightingPreset.LegacyFlat2 => A.LightRigValues.LegacyFlat2,
                LightingPreset.LegacyFlat3 => A.LightRigValues.LegacyFlat3,
                LightingPreset.LegacyFlat4 => A.LightRigValues.LegacyFlat4,
                LightingPreset.LegacyNormal1 => A.LightRigValues.LegacyNormal1,
                LightingPreset.LegacyNormal2 => A.LightRigValues.LegacyNormal2,
                LightingPreset.LegacyNormal3 => A.LightRigValues.LegacyNormal3,
                LightingPreset.LegacyNormal4 => A.LightRigValues.LegacyNormal4,
                LightingPreset.LegacyHarsh1 => A.LightRigValues.LegacyHarsh1,
                LightingPreset.LegacyHarsh2 => A.LightRigValues.LegacyHarsh2,
                LightingPreset.LegacyHarsh3 => A.LightRigValues.LegacyHarsh3,
                LightingPreset.LegacyHarsh4 => A.LightRigValues.LegacyHarsh4,
                LightingPreset.ThreePoints => A.LightRigValues.ThreePoints,
                LightingPreset.Balanced => A.LightRigValues.Balanced,
                LightingPreset.Soft => A.LightRigValues.Soft,
                LightingPreset.Harsh => A.LightRigValues.Harsh,
                LightingPreset.Flood => A.LightRigValues.Flood,
                LightingPreset.Contrasting => A.LightRigValues.Contrasting,
                LightingPreset.Morning => A.LightRigValues.Morning,
                LightingPreset.Sunrise => A.LightRigValues.Sunrise,
                LightingPreset.Sunset => A.LightRigValues.Sunset,
                LightingPreset.Chilly => A.LightRigValues.Chilly,
                LightingPreset.Freezing => A.LightRigValues.Freezing,
                LightingPreset.Flat => A.LightRigValues.Flat,
                LightingPreset.TwoPoints => A.LightRigValues.TwoPoints,
                LightingPreset.Glow => A.LightRigValues.Glow,
                LightingPreset.BrightRoom => A.LightRigValues.BrightRoom,
                _ => A.LightRigValues.Soft
            };
        }
        private static A.PresetShadowValues mapShadowPreset(ShadowPreset? value)
        {
            if (value is null)
                return A.PresetShadowValues.FrontBottomShadow;

            return value.Value switch
            {
                ShadowPreset.TopLeftDropShadow => A.PresetShadowValues.TopLeftDropShadow,
                ShadowPreset.TopRightDropShadow => A.PresetShadowValues.TopRightDropShadow,
                ShadowPreset.BackLeftPerspectiveShadow => A.PresetShadowValues.BackLeftPerspectiveShadow,
                ShadowPreset.BackRightPerspectiveShadow => A.PresetShadowValues.BackRightPerspectiveShadow,
                ShadowPreset.BottomLeftDropShadow => A.PresetShadowValues.BottomLeftDropShadow,
                ShadowPreset.BottomRightDropShadow => A.PresetShadowValues.BottomRightDropShadow,
                ShadowPreset.FrontLeftPerspectiveShadow => A.PresetShadowValues.FrontLeftPerspectiveShadow,
                ShadowPreset.FrontRightPerspectiveShadow => A.PresetShadowValues.FrontRightPerspectiveShadow,
                ShadowPreset.TopLeftSmallDropShadow => A.PresetShadowValues.TopLeftSmallDropShadow,
                ShadowPreset.TopLeftLargeDropShadow => A.PresetShadowValues.TopLeftLargeDropShadow,
                ShadowPreset.BackLeftLongPerspectiveShadow => A.PresetShadowValues.BackLeftLongPerspectiveShadow,
                ShadowPreset.BackRightLongPerspectiveShadow => A.PresetShadowValues.BackRightLongPerspectiveShadow,
                ShadowPreset.TopLeftDoubleDropShadow => A.PresetShadowValues.TopLeftDoubleDropShadow,
                ShadowPreset.BottomRightSmallDropShadow => A.PresetShadowValues.BottomRightSmallDropShadow,
                ShadowPreset.FrontLeftLongPerspectiveShadow => A.PresetShadowValues.FrontLeftLongPerspectiveShadow,
                ShadowPreset.FrontRightLongPerspectiveShadow => A.PresetShadowValues.FrontRightLongPerspectiveShadow,
                ShadowPreset.ThreeDimensionalOuterBoxShadow => A.PresetShadowValues.ThreeDimensionalOuterBoxShadow,
                ShadowPreset.ThreeDimensionalInnerBoxShadow => A.PresetShadowValues.ThreeDimensionalInnerBoxShadow,
                ShadowPreset.BackCenterPerspectiveShadow => A.PresetShadowValues.BackCenterPerspectiveShadow,
                ShadowPreset.FrontBottomShadow => A.PresetShadowValues.FrontBottomShadow,
                _ => A.PresetShadowValues.FrontBottomShadow,
            };
        }
        private static A.TextAlignmentTypeValues mapTextAlign(HorizontalAlign value)
        {
            return value switch
            {
                HorizontalAlign.Left => A.TextAlignmentTypeValues.Left,
                HorizontalAlign.Center => A.TextAlignmentTypeValues.Center,
                HorizontalAlign.Right => A.TextAlignmentTypeValues.Right,
                HorizontalAlign.Justify => A.TextAlignmentTypeValues.Justified,
                _ => A.TextAlignmentTypeValues.Left
            };
        }
        private static A.TextAnchoringTypeValues mapTextAnchor(VerticalAlign value)
        {
            return value switch
            {
                VerticalAlign.Top => A.TextAnchoringTypeValues.Top,
                VerticalAlign.Center => A.TextAnchoringTypeValues.Center,
                VerticalAlign.Bottom => A.TextAnchoringTypeValues.Bottom,
                _ => A.TextAnchoringTypeValues.Center
            };
        }
        private static A.TextWrappingValues mapTextWrap(TextWrapPreset value)
        {
            return value switch
            {
                TextWrapPreset.None => A.TextWrappingValues.None,
                TextWrapPreset.Square => A.TextWrappingValues.Square,
                _ => A.TextWrappingValues.Square
            };
        }
        private static A.TextUnderlineValues mapUnderline(TextUnderlineStyle value)
        {
            return value switch
            {
                TextUnderlineStyle.None => A.TextUnderlineValues.None,
                TextUnderlineStyle.Single => A.TextUnderlineValues.Single,
                TextUnderlineStyle.Double => A.TextUnderlineValues.Double,
                _ => A.TextUnderlineValues.None
            };
        }
        private static A.TextStrikeValues mapStrike(TextStrikeStyle value)
        {
            return value switch
            {
                TextStrikeStyle.None => A.TextStrikeValues.NoStrike,
                TextStrikeStyle.Single => A.TextStrikeValues.SingleStrike,
                TextStrikeStyle.Double => A.TextStrikeValues.DoubleStrike,
                _ => A.TextStrikeValues.NoStrike
            };
        }
        private static A.TextCapsValues mapCaps(TextCapsStyle value)
        {
            return value switch
            {
                TextCapsStyle.None => A.TextCapsValues.None,
                TextCapsStyle.Small => A.TextCapsValues.Small,
                TextCapsStyle.All => A.TextCapsValues.All,
                _ => A.TextCapsValues.None
            };
        }

    }

    class ChartDefinition : IValueLabelsContainer
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
        public ChartDefinition GetChartDefinition()
        {
            return this;
        }
    }

    class SeriesDefinition : IValueLabelsContainer
    {
        public TitleDefinition Title { get; set; }
        public ColorA ColorA { get { return Format?.FillDefinition?.SolidFillDefinition?.Color; } }
        public string Color { get { return ColorA?.Color; } }
        public int? ColorTransparency { get { return ColorA?.Transparency; } }
        public List<PointDefinition> Points { get; set; } = [];
        public FormatDefinition Format { get; set; }
        public MarkerDefinition Marker { get; set; }
        public ValueLabelDefinition ValueLabels { get; set; } //<value-labels>
        public ChartDefinition ChartDefinition { get; set; }
        public SeriesDefinition()
        {

        }
        public  /*Ctor*/                SeriesDefinition(ChartDefinition chartDefinition)
        {
            ChartDefinition = chartDefinition;
        }
        public ChartDefinition GetChartDefinition()
        {
            return ChartDefinition;
        }
    }

    class PointDefinition : IValueLabelsContainer
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
        public PointDefinition()
        {

        }
        public  /*Ctor*/                PointDefinition(SeriesDefinition seriesDefinition)
        {
            SeriesDefinition = seriesDefinition;
        }
        public ChartDefinition GetChartDefinition()
        {
            return SeriesDefinition?.GetChartDefinition();
        }
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


            if (!string.IsNullOrEmpty(transparencyStr) && int.TryParse(transparencyStr, out int transparencyValue))
            {
                if (transparencyValue < 0)
                    transparency = 0;
                else if (transparencyValue > 100)
                    transparency = 100;
                else
                    transparency = transparencyValue;
            }

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

    internal interface IValueLabelsContainer
    {
        ValueLabelDefinition ValueLabels { get; }
        ChartDefinition GetChartDefinition();
    }

}


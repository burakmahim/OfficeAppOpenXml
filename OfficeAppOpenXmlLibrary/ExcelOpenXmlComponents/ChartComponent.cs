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
using System;

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

                    XElement? effectsNode = formatNode.Element("effects");
                    if (effectsNode != null)
                    {
                        chartDefinition.CategoryAxis.AxisLineFormat.EffectsDefinition = new EffectsDefinition();

                        XElement? shadowNode = effectsNode.Element("shadow");
                        if (shadowNode != null)
                        {
                            chartDefinition.CategoryAxis.AxisLineFormat.EffectsDefinition.ShadowDefinition = new ShadowDefinition();

                            Nullable<ShadowType> shadowType = XElementAttributeGetter.AsEnum<ShadowType>(shadowNode, "type");
                            if (shadowType.HasValue)
                                chartDefinition.CategoryAxis.AxisLineFormat.EffectsDefinition.ShadowDefinition.Type = shadowType.Value;

                            Nullable<ShadowPreset> shadowPreset = XElementAttributeGetter.AsEnum<ShadowPreset>(shadowNode, "preset");
                            if (shadowPreset.HasValue)
                                chartDefinition.CategoryAxis.AxisLineFormat.EffectsDefinition.ShadowDefinition.Preset = shadowPreset.Value;

                            string? blurStr = shadowNode.Attribute("blur-radius")?.Value;
                            if (!string.IsNullOrEmpty(blurStr))
                                chartDefinition.CategoryAxis.AxisLineFormat.EffectsDefinition.ShadowDefinition.BlurRadius = Dimension.Parse(blurStr);

                            string? distanceStr = shadowNode.Attribute("distance")?.Value;
                            if (!string.IsNullOrEmpty(distanceStr))
                                chartDefinition.CategoryAxis.AxisLineFormat.EffectsDefinition.ShadowDefinition.Distance = Dimension.Parse(distanceStr);

                            if (XElementAttributeGetter.AsInt32(shadowNode, "angle", out int angle))
                                chartDefinition.CategoryAxis.AxisLineFormat.EffectsDefinition.ShadowDefinition.Angle = angle;

                            ColorA shadowColor = ColorA.Parse(shadowNode, "color", "transparency");
                            if (shadowColor != null)
                                chartDefinition.CategoryAxis.AxisLineFormat.EffectsDefinition.ShadowDefinition.Color = shadowColor;
                        }

                        XElement? glowNode = effectsNode.Element("glow");
                        if (glowNode != null)
                        {
                            chartDefinition.CategoryAxis.AxisLineFormat.EffectsDefinition.GlowDefinition = new GlowDefinition();

                            ColorA shadowColor = ColorA.Parse(glowNode, "color", "transparency");
                            if (shadowColor != null)
                                chartDefinition.CategoryAxis.AxisLineFormat.EffectsDefinition.GlowDefinition.Color = shadowColor;

                        }
                    }
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

                XElement? formatNode = valueAxisNode.Element("format");
                if (formatNode != null)
                {
                    chartDefinition.ValueAxis.AxisLineFormat = new FormatDefinition();

                    XElement? fillNode = formatNode.Element("fill");
                    if (fillNode != null)
                    {
                        chartDefinition.ValueAxis.AxisLineFormat.FillDefinition = new FillDefinition();

                        XElement? solidNode = fillNode.Element("solid");
                        if (solidNode != null)
                        {
                            chartDefinition.ValueAxis.AxisLineFormat.FillDefinition.SolidFillDefinition = new SolidDefinition();

                            ColorA color = ColorA.Parse(solidNode, "color", "transparency");
                            if (color != null)
                                chartDefinition.ValueAxis.AxisLineFormat.FillDefinition.SolidFillDefinition.Color = color;
                        }

                        XElement? patternNode = fillNode.Element("pattern");
                        if (patternNode != null)
                        {
                            chartDefinition.ValueAxis.AxisLineFormat.FillDefinition.PatternFillDefinition = new PatternDefinition();

                            Nullable<PatternFillPreset> patternType = XElementAttributeGetter.AsEnum<PatternFillPreset>(patternNode, "preset");
                            if (patternType.HasValue)
                                chartDefinition.ValueAxis.AxisLineFormat.FillDefinition.PatternFillDefinition.Preset = patternType.Value;

                            ColorA foregroundColor = ColorA.Parse(patternNode, "foreground-color", "foreground-transparency");
                            if (foregroundColor != null)
                                chartDefinition.ValueAxis.AxisLineFormat.FillDefinition.PatternFillDefinition.ForegroundColor = foregroundColor;

                            ColorA backgroundColor = ColorA.Parse(patternNode, "background-color", "background-transparency");
                            if (backgroundColor != null)
                                chartDefinition.ValueAxis.AxisLineFormat.FillDefinition.PatternFillDefinition.BackgroundColor = backgroundColor;
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

                            chartDefinition.ValueAxis.AxisLineFormat.FillDefinition.GradientFillDefinition = gradientDef;
                        }

                    }

                    XElement? lineNode = formatNode.Element("line");
                    if (lineNode != null)
                    {
                        chartDefinition.ValueAxis.AxisLineFormat.LineDefinition = new LineDefinition();

                        XElementAttributeGetter.AsBool(lineNode, "visible", out bool visibleValue);
                        chartDefinition.ValueAxis.AxisLineFormat.LineDefinition.Visible = visibleValue;

                        Nullable<DashPreset> dashPreset = XElementAttributeGetter.AsEnum<DashPreset>(lineNode, "dash");
                        if (dashPreset.HasValue)
                            chartDefinition.ValueAxis.AxisLineFormat.LineDefinition.DashPreset = dashPreset.Value;

                        Nullable<CompoundLinePreset> compoundLine = XElementAttributeGetter.AsEnum<CompoundLinePreset>(lineNode, "compound");
                        if (compoundLine.HasValue)
                            chartDefinition.ValueAxis.AxisLineFormat.LineDefinition.CompoundPreset = compoundLine.Value;

                        Nullable<LineCapPreset> lineCap = XElementAttributeGetter.AsEnum<LineCapPreset>(lineNode, "cap");
                        if (lineCap.HasValue)
                            chartDefinition.ValueAxis.AxisLineFormat.LineDefinition.CapPreset = lineCap.Value;

                        Nullable<LineJoinPreset> join = XElementAttributeGetter.AsEnum<LineJoinPreset>(lineNode, "join");
                        if (join.HasValue)
                            chartDefinition.ValueAxis.AxisLineFormat.LineDefinition.JoinPreset = join.Value;

                        Nullable<LineEndPreset> beginArrowType = XElementAttributeGetter.AsEnum<LineEndPreset>(lineNode, "begin-arrow-type");
                        if (beginArrowType.HasValue)
                            chartDefinition.ValueAxis.AxisLineFormat.LineDefinition.Begin_Arrow_Type = beginArrowType.Value;

                        Nullable<LineEndWidthPreset> beginArrowWidth = XElementAttributeGetter.AsEnum<LineEndWidthPreset>(lineNode, "begin-arrow-width");
                        if (beginArrowWidth.HasValue)
                            chartDefinition.ValueAxis.AxisLineFormat.LineDefinition.Begin_Arrow_Width = beginArrowWidth.Value;

                        Nullable<LineEndLengthPreset> beginArrowLength = XElementAttributeGetter.AsEnum<LineEndLengthPreset>(lineNode, "begin-arrow-length");
                        if (beginArrowLength.HasValue)
                            chartDefinition.ValueAxis.AxisLineFormat.LineDefinition.Begin_Arrow_Length = beginArrowLength.Value;

                        Nullable<LineEndPreset> endArrowType = XElementAttributeGetter.AsEnum<LineEndPreset>(lineNode, "end-arrow-type");
                        if (endArrowType.HasValue)
                            chartDefinition.ValueAxis.AxisLineFormat.LineDefinition.End_Arrow_Type = endArrowType.Value;

                        Nullable<LineEndWidthPreset> endArrowWidth = XElementAttributeGetter.AsEnum<LineEndWidthPreset>(lineNode, "end-arrow-width");
                        if (endArrowWidth.HasValue)
                            chartDefinition.ValueAxis.AxisLineFormat.LineDefinition.End_Arrow_Width = endArrowWidth.Value;

                        Nullable<LineEndLengthPreset> endArrowLength = XElementAttributeGetter.AsEnum<LineEndLengthPreset>(lineNode, "end-arrow-length");
                        if (endArrowLength.HasValue)
                            chartDefinition.ValueAxis.AxisLineFormat.LineDefinition.End_Arrow_Length = endArrowLength.Value;

                        if (XElementAttributeGetter.AsInt32(lineNode, "join-miter-limit", out int joinMiterLimitValue))
                        {
                            int clampedValue = clamp(joinMiterLimitValue, 1, 500);
                            chartDefinition.ValueAxis.AxisLineFormat.LineDefinition.JoinMiterLimit = clampedValue;
                        }

                    }

                    XElement? effectsNode = formatNode.Element("effects");
                    if (effectsNode != null)
                    {
                        chartDefinition.ValueAxis.AxisLineFormat.EffectsDefinition = new EffectsDefinition();

                        XElement? shadowNode = effectsNode.Element("shadow");
                        if (shadowNode != null)
                        {
                            chartDefinition.ValueAxis.AxisLineFormat.EffectsDefinition.ShadowDefinition = new ShadowDefinition();

                            Nullable<ShadowType> shadowType = XElementAttributeGetter.AsEnum<ShadowType>(shadowNode, "type");
                            if (shadowType.HasValue)
                                chartDefinition.ValueAxis.AxisLineFormat.EffectsDefinition.ShadowDefinition.Type = shadowType.Value;

                            Nullable<ShadowPreset> shadowPreset = XElementAttributeGetter.AsEnum<ShadowPreset>(shadowNode, "preset");
                            if (shadowPreset.HasValue)
                                chartDefinition.ValueAxis.AxisLineFormat.EffectsDefinition.ShadowDefinition.Preset = shadowPreset.Value;

                            string? blurStr = shadowNode.Attribute("blur-radius")?.Value;
                            if (!string.IsNullOrEmpty(blurStr))
                                chartDefinition.ValueAxis.AxisLineFormat.EffectsDefinition.ShadowDefinition.BlurRadius = Dimension.Parse(blurStr);

                            string? distanceStr = shadowNode.Attribute("distance")?.Value;
                            if (!string.IsNullOrEmpty(distanceStr))
                                chartDefinition.ValueAxis.AxisLineFormat.EffectsDefinition.ShadowDefinition.Distance = Dimension.Parse(distanceStr);

                            if (XElementAttributeGetter.AsInt32(shadowNode, "angle", out int angle))
                                chartDefinition.ValueAxis.AxisLineFormat.EffectsDefinition.ShadowDefinition.Angle = angle;

                            ColorA shadowColor = ColorA.Parse(shadowNode, "color", "transparency");
                            if (shadowColor != null)
                                chartDefinition.ValueAxis.AxisLineFormat.EffectsDefinition.ShadowDefinition.Color = shadowColor;
                        }

                        XElement? glowNode = effectsNode.Element("glow");
                        if (glowNode != null)
                        {
                            chartDefinition.ValueAxis.AxisLineFormat.EffectsDefinition.GlowDefinition = new GlowDefinition();

                            ColorA shadowColor = ColorA.Parse(glowNode, "color", "transparency");
                            if (shadowColor != null)
                                chartDefinition.ValueAxis.AxisLineFormat.EffectsDefinition.GlowDefinition.Color = shadowColor;

                        }
                    }
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


                XElement? formatNode = legendNode.Element("format");
                if (formatNode != null)
                {
                    chartDefinition.Legend.BoxFormat = new FormatDefinition();

                    XElement? fillNode = formatNode.Element("fill");
                    if (fillNode != null)
                    {
                        chartDefinition.Legend.BoxFormat.FillDefinition = new FillDefinition();

                        XElement? solidNode = fillNode.Element("solid");
                        if (solidNode != null)
                        {
                            chartDefinition.Legend.BoxFormat.FillDefinition.SolidFillDefinition = new SolidDefinition();

                            ColorA color = ColorA.Parse(solidNode, "color", "transparency");
                            if (color != null)
                                chartDefinition.Legend.BoxFormat.FillDefinition.SolidFillDefinition.Color = color;
                        }

                        XElement? patternNode = fillNode.Element("pattern");
                        if (patternNode != null)
                        {
                            chartDefinition.Legend.BoxFormat.FillDefinition.PatternFillDefinition = new PatternDefinition();

                            Nullable<PatternFillPreset> patternType = XElementAttributeGetter.AsEnum<PatternFillPreset>(patternNode, "preset");
                            if (patternType.HasValue)
                                chartDefinition.Legend.BoxFormat.FillDefinition.PatternFillDefinition.Preset = patternType.Value;

                            ColorA foregroundColor = ColorA.Parse(patternNode, "foreground-color", "foreground-transparency");
                            if (foregroundColor != null)
                                chartDefinition.Legend.BoxFormat.FillDefinition.PatternFillDefinition.ForegroundColor = foregroundColor;

                            ColorA backgroundColor = ColorA.Parse(patternNode, "background-color", "background-transparency");
                            if (backgroundColor != null)
                                chartDefinition.Legend.BoxFormat.FillDefinition.PatternFillDefinition.BackgroundColor = backgroundColor;
                        }
                    }

                    XElement? lineNode = formatNode.Element("line");
                    if (lineNode != null)
                    {
                        chartDefinition.Legend.BoxFormat.LineDefinition = new LineDefinition();

                        XElementAttributeGetter.AsBool(lineNode, "visible", out bool visibleValue);
                        chartDefinition.Legend.BoxFormat.LineDefinition.Visible = visibleValue;

                        Nullable<DashPreset> dashPreset = XElementAttributeGetter.AsEnum<DashPreset>(lineNode, "dash");
                        if (dashPreset.HasValue)
                            chartDefinition.Legend.BoxFormat.LineDefinition.DashPreset = dashPreset.Value;

                        Nullable<CompoundLinePreset> compoundLine = XElementAttributeGetter.AsEnum<CompoundLinePreset>(lineNode, "compound");
                        if (compoundLine.HasValue)
                            chartDefinition.Legend.BoxFormat.LineDefinition.CompoundPreset = compoundLine.Value;
                    }

                    XElement? effectsNode = formatNode.Element("effects");
                    if (effectsNode != null)
                    {
                        chartDefinition.Legend.BoxFormat.EffectsDefinition = new EffectsDefinition();

                        XElement? shadowNode = effectsNode.Element("shadow");
                        if (shadowNode != null)
                        {
                            chartDefinition.Legend.BoxFormat.EffectsDefinition.ShadowDefinition = new ShadowDefinition();

                            Nullable<ShadowType> shadowType = XElementAttributeGetter.AsEnum<ShadowType>(shadowNode, "type");
                            if (shadowType.HasValue)
                                chartDefinition.Legend.BoxFormat.EffectsDefinition.ShadowDefinition.Type = shadowType.Value;

                            ColorA shadowColor = ColorA.Parse(shadowNode, "color", "transparency");
                            if (shadowColor != null)
                                chartDefinition.Legend.BoxFormat.EffectsDefinition.ShadowDefinition.Color = shadowColor;
                        }

                        XElement? glowNode = effectsNode.Element("glow");
                        if (glowNode != null)
                        {
                            chartDefinition.Legend.BoxFormat.EffectsDefinition.GlowDefinition = new GlowDefinition();

                            ColorA glowColor = ColorA.Parse(glowNode, "color", "transparency");
                            if (glowColor != null)
                                chartDefinition.Legend.BoxFormat.EffectsDefinition.GlowDefinition.Color = glowColor;
                        }
                    }
                }

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

                addBarSeries(chartDefinition, chartNode, bar3DChart);

                //addDataLabels(bar3DChart, chartDefinition);

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

                addBarSeries(chartDefinition, chartNode, barChart);

                if (chartDefinition.Overlap.HasValue)
                    barChart.Append(new Overlap() { Val = new SByteValue((sbyte)chartDefinition.Overlap.Value) });

                if (chartDefinition.GapWidth.HasValue)
                    barChart.Append(new GapWidth() { Val = new UInt16Value((ushort)chartDefinition.GapWidth.Value) });

                //addDataLabels(barChart, chartDefinition);

                plotArea.Append(barChart);
            }

            addCategoryAxis(chartDefinition, chartNode, plotArea, categoryAxisId, valueAxisId, categoryAxisPosition);
            addValueAxis(chartDefinition, chartNode, plotArea, categoryAxisId, valueAxisId, valueAxisPosition);
            //addLayout(chartDefinition.Layout, plotArea);----------
            //addWallsAndFloor(plotArea, chartDefinition.ThreeDViewDefinition);

            return GetChartCommon(chartDefinition, chartNode, plotArea);
        }
        private static C.Chart getLineChart(ChartDefinition chartDefinition, XElement chartNode)
        {
            uint catId = getSafeId();
            uint valId = getSafeId();

            C.PlotArea plotArea = new();
            plotArea.Append(new Layout());

            bool use3D = chartDefinition.ThreeDView is not null;

            if (use3D)
            {
                Line3DChart line3DChart = new Line3DChart()
                {
                    Grouping = new Grouping() { Val = mapGrouping(chartDefinition.GroupingType ?? GroupingType.Standard) },
                    VaryColors = new VaryColors() { Val = chartDefinition.VaryColors ?? false }
                };

                addAxisIds(line3DChart, catId, valId);
                addLineSeries(chartDefinition, chartNode, line3DChart);
                //addDataLabels(line3DChart, chartDefinition);

                plotArea.Append(line3DChart);

                addCategoryAxis(chartDefinition, chartNode, plotArea, catId, valId);
                addValueAxis(chartDefinition, chartNode, plotArea, catId, valId);
                //addLayout(chartDefinition.Layout, plotArea);

                return GetChartCommon(chartDefinition, chartNode, plotArea);
            }
            else
            {
                LineChart lineChart = new()
                {
                    Grouping = new Grouping() { Val = mapGrouping(chartDefinition.GroupingType ?? GroupingType.Standard) },
                    VaryColors = new VaryColors() { Val = chartDefinition.VaryColors ?? false }
                };

                addAxisIds(lineChart, catId, valId);

                addLineSeries(chartDefinition, chartNode, lineChart);
                //addDataLabels(lineChart, chartDefinition);

                plotArea.Append(lineChart);

                addCategoryAxis(chartDefinition, chartNode, plotArea, catId, valId);
                addValueAxis(chartDefinition, chartNode, plotArea, catId, valId);
                //addLayout(chartDefinition.Layout, plotArea);

                return GetChartCommon(chartDefinition, chartNode, plotArea);
            }
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

                    //if (applyTextFormat(textProperties, legendDefinition.LegendEntryTextFormat))
                    //    legend.Append(textProperties);
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

                        //if (applyTextFormat(textProperties, textFormatDefinition))
                        //    legend.Append(legendEntry);
                    }

                    index++;
                }

                if (legendDefinition.BoxFormat is not null)
                {
                    applyFormat(legend, legendDefinition.BoxFormat);
                }

                chart.Append(legend);
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
        private static void addBarSeries(ChartDefinition chartDefinition, XElement chartNode, OpenXmlCompositeElement barChart)
        {
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

                //applyPointLevelStyling(seriesDefinition, series, useMarker: false);

                series.Append(catAxisData);
                series.Append(values);

                //(chartDefinition, seriesDefinition, series);
                //addDataLabels(series, seriesDefinition);

                barChart.Append(series);

                seriesIndex++;
            }
        }
        private static void addLineSeries(ChartDefinition chartDefinition, XElement chartNode, OpenXmlCompositeElement lineChart)
        {
            uint idx = 0;

            foreach (XElement seriesNode in chartNode.Elements("series"))
            {
                C.LineChartSeries series = new(
                    new C.Index() { Val = idx },
                    new C.Order() { Val = idx },
                    new C.SeriesText(new C.NumericValue(seriesNode.Attribute("name")?.Value ?? ""))
                );

                CategoryAxisData catAxisData = new CategoryAxisData();
                C.Values values = new C.Values();

                List<XElement> points = seriesNode.Elements("point").ToList();
                uint pointCount = (uint)points.Count;

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

            ValueAxis valAx = new(
                new C.AxisId() { Val = valueAxisId },
                new Scaling(new Orientation() { Val = C.OrientationValues.MinMax }),
                new AxisPosition() { Val = position },
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

            //if (addMajorGridLines)
            //    addMajorGridlines(valAx, targetMajorGridlinesFormat);

            valAx.InsertAt(new Delete() { Val = !showAxis }, 2);

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

            //if (effectsDefinition.Format3dDefinition is not null)
            //{
            //    applyFormat3D(chartShapeProperties, effectsDefinition.Format3dDefinition);
            //}

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
        private static string ToHexColorCode(string colorOrName)
        {
            if (string.IsNullOrWhiteSpace(colorOrName))
                return null;

            colorOrName = colorOrName.Trim().ToLower();

            if (colorOrName.StartsWith("#"))
                return colorOrName.Substring(1);

            return colorOrName switch
            {
                "black" => "000000",
                "white" => "FFFFFF",
                "red" => "FF0000",
                "green" => "00FF00",
                "blue" => "0000FF",
                "yellow" => "FFFF00",
                "cyan" => "00FFFF",
                "magenta" => "FF00FF",
                "gray" => "808080",
                "grey" => "808080",
                _ => null
            };
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



}



// -----------------------------------------------------------------------------
// Excel.ChartDefinition.cs
// Excel uyumlu ChartDefinition + yardımcıları
// Word bağımlılıkları kaldırıldı; ExcelDocument/ExcelXMLAttributeGetter eklendi.
// -----------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents
{
    // -------------------------------------------------------------------------
    // Excel tarafı attribute/tag yardımcıları (WordDocument yerine)
    // -------------------------------------------------------------------------
    internal static class ExcelDocument
    {
        internal static class Tag
        {
            public const string category_axis = "category_axis";
            public const string value_axis = "value_axis";
            public const string legend = "legend";
            public const string grid = "grid";
            public const string value_labels = "value_labels";
            public const string series = "series";
            public const string layout = "layout";
            public const string view_3d = "view_3d";
            public const string plot_area = "plot_area";
            public const string format = "format";
            public const string text_format = "text_format";
            public const string marker = "marker";
            public const string title = "title";
            public const string series_format = "series_format";
            public const string data_table = "data_table";
            public const string horizontal_format = "horizontal_format";
            public const string vertical_format = "vertical_format";
            public const string stop = "stop";
            public const string solid = "solid";
            public const string pattern = "pattern";
            public const string gradient = "gradient";
            public const string effects = "effects";
            public const string shadow = "shadow";
            public const string glow = "glow";
            public const string soft_edges = "soft_edges";
            public const string format_3D = "format_3d";
            public const string format_3d = "format_3d";
            public const string reflection = "reflection";
            public const string font = "font";
            public const string fill = "fill";
            public const string line = "line";
            public const string text = "text";
            public const string floor_format = "floor_format";
            public const string back_wall_format = "back_wall_format";
            public const string side_wall_format = "side_wall_format";
            public const string bevel = "bevel";
            public const string lighting = "lighting";
            public const string point = "point";
        }

        internal enum Attribute
        {
            // genel
            type, title, width, height, display_blanks_as, grouping_type, auto_title, plot_visible_only, rounded_corners, vary_colors, show_data_labels_over_maximum, create_new_paragraph,
            show_category_axis, show_value_axis, first_slice_angle, doughnut_hole_size, bubble_scale, bubble_3d, overlap, gap_width, text, preset, direction,

            // scatter/radar
            scatter_style, radar_style,

            // legend/grid/value labels
            show, position, show_horizontal, show_vertical, show_legend_key, show_category_name, show_series_name, show_percent, show_bubble_size,

            // layout
            x, y, x_mode, y_mode, target,

            // view_3d
            shape, rotation_x, rotation_y, perspective, depth_percent, height_percent, gap_depth, right_angle_axes, show_floor, show_back_wall, show_side_wall,

            // series/point
            color, color_transparency, marker_type, marker_size, category, value, size,

            // axis
            min, max, major_unit,

            // text-format
            align_horizontal, align_vertical, wrap, margin_top, margin_right, margin_bottom, margin_left, rotate,

            // fill/pattern/gradient
            foreground_color, background_color, foreground_transparency, background_transparency, angle, scaled,

            // gradient stop
            transparency,

            // line
            dash, compound, cap, join, begin_arrow_type, begin_arrow_width, begin_arrow_length, end_arrow_type, end_arrow_width, end_arrow_length, join_miter_limit, visible,

            // effects
            blur_radius, distance, start_transparency, end_transparency, start_position, end_position,

            // data-table (EKLENMELİ)
            show_horizontal_border, show_vertical_border, show_outline_border,

            // format-3d
            material, top_width, top_height, bottom_width, bottom_height, top_preset, bottom_preset,

            // font
            family, underline, strike, caps, kerning, spacing, bold, italic,

            // scatter values
            x_value, y_value // (opsiyonel alias)
        }
    }

    internal static class ExcelXMLAttributeGetter
    {
        private static string AttrName(ExcelDocument.Attribute a) => a switch
        {
            // direct string matches
            _ => a.ToString().Replace("_value", "").Replace("__", "_")
        };

        internal static string AsString(XmlNode node, ExcelDocument.Attribute attr)
            => node?.Attributes?[AttrName(attr)]?.Value;

        internal static string AsStringSafe(XmlNode node, ExcelDocument.Attribute attr)
            => AsString(node, attr) ?? string.Empty;

        internal static bool AsBool(XmlNode node, ExcelDocument.Attribute attr, out bool value)
        {
            value = false;
            var s = AsString(node, attr);
            if (string.IsNullOrWhiteSpace(s)) return false;
            if (bool.TryParse(s, out var b)) { value = b; return true; }
            // Excel şemasında 0/1 de gelebilir
            if (s == "0") { value = false; return true; }
            if (s == "1") { value = true; return true; }
            return false;
        }

        internal static bool AsBoolSafe(XmlNode node, ExcelDocument.Attribute attr)
            => AsBool(node, attr, out var b) && b;

        internal static bool AsInt32(XmlNode node, ExcelDocument.Attribute attr, out int value)
        {
            value = 0;
            var s = AsString(node, attr);
            return int.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out value);
        }

        internal static bool AsDouble(XmlNode node, ExcelDocument.Attribute attr, out double value)
        {
            value = 0;
            var s = AsString(node, attr);
            return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out value);
        }

        internal static TEnum? AsEnum<TEnum>(XmlNode node, ExcelDocument.Attribute attr) where TEnum : struct
        {
            var s = AsString(node, attr);
            if (string.IsNullOrWhiteSpace(s)) return null;
            // kebab-case / snake-case olabilir → normalize
            string norm = s.Replace("-", "").Replace("_", "").Trim();
            foreach (var name in Enum.GetNames(typeof(TEnum)))
            {
                if (string.Equals(name.Replace("_", "").ToLowerInvariant(), norm.ToLowerInvariant(), StringComparison.Ordinal))
                {
                    if (Enum.TryParse<TEnum>(name, true, out var val)) return val;
                }
            }
            if (Enum.TryParse<TEnum>(s, true, out var val2)) return val2;
            return null;
        }
    }


    public static class ChartDefinitionHelper
    {
        internal static int? ParseInt32WithMinMax(this XmlNode node, ExcelDocument.Attribute attr, int? min = null, int? max = null)
        {
            if (node is null || !ExcelXMLAttributeGetter.AsInt32(node, attr, out int result))
                return null;

            if (min is not null && result < min)
                return null;

            if (max is not null && result > max)
                return null;

            return result;
        }
        internal static double? ParseDoubleWithMinMax(this XmlNode node, ExcelDocument.Attribute attr, double? min = null, double? max = null)
        {
            if (node is null || !ExcelXMLAttributeGetter.AsDouble(node, attr, out double result))
                return null;

            if (min is not null && result < min)
                return null;

            if (max is not null && result > max)
                return null;

            return result;
        }
        internal static bool? ParseBool(this XmlNode node, ExcelDocument.Attribute attr)
        {
            if (node is not null && ExcelXMLAttributeGetter.AsBool(node, attr, out bool result))
                return result;

            return null;
        }

        internal static void ApplyTitleShortcut(this ITitleOwner owner, XmlNode node, ExcelDocument.Attribute titleAttribute = ExcelDocument.Attribute.title)
        {
            if (owner is null || node is null)
                return;

            owner.Title ??= new();

            if (!string.IsNullOrWhiteSpace(owner.Title.Text))
                return;

            string simpleTitle = ExcelXMLAttributeGetter.AsString(node, titleAttribute);

            if (!string.IsNullOrWhiteSpace(simpleTitle))
                owner.Title.Text = simpleTitle;
        }
        internal static void ApplyMarkerShortcut(this IMarkerOwner owner, XmlNode node, ExcelDocument.Attribute markerSizeAttr = ExcelDocument.Attribute.marker_size, ExcelDocument.Attribute markerTypeAttr = ExcelDocument.Attribute.marker_type)
        {
            if (owner is null || node is null)
                return;

            owner.Marker ??= new();

            if (owner.Marker.Type is null)
                owner.Marker.Type = ExcelXMLAttributeGetter.AsEnum<MarkerType>(node, markerTypeAttr);

            if (owner.Marker.Size is null)
                owner.Marker.Size = node.ParseInt32WithMinMax(markerSizeAttr, 2, 72);
        }
    }

    class ViewChartDefinition : ChartDefinition
    {
        private IViewChartDefinition viewChartDefinition { get; set; }
        internal override bool Analyse(XmlNode chartNode, ExcelDocumentAssets assets)
        {
            viewChartDefinition = assets?.ChartLoadingProvider?.LoadViewChart(chartNode);

            if (viewChartDefinition is null)
                return false;

            if (!string.IsNullOrWhiteSpace(viewChartDefinition.Title))
            {
                Title = new()
                {
                    Text = viewChartDefinition.Title
                };
            }

            if (viewChartDefinition.Series != null && viewChartDefinition.Series.Any())
            {
                foreach (var series in viewChartDefinition.Series)
                {
                    SeriesDefinition seriesDefinition = new(this)
                    {
                        Title = string.IsNullOrWhiteSpace(series.Title) ? null : new TitleDefinition { Text = series.Title }
                    };

                    string seriesColor = ColorUtil.ToHex(series.Color);

                    if (!string.IsNullOrWhiteSpace(seriesColor))
                    {
                        seriesDefinition.Format = new()
                        {
                            FillDefinition = new()
                            {
                                SolidFillDefinition = new()
                                {
                                    Color = new ColorA(seriesColor, null)
                                }
                            }
                        };

                        if (IsLineBased)
                        {
                            seriesDefinition.Format.LineDefinition = new()
                            {
                                Visible = true,
                                DashPreset = DashPreset.Solid,
                                SolidLineDefinition = new()
                                {
                                    Color = new ColorA(seriesColor, null)
                                }
                            };
                        }
                    }

                    if (series.Points != null && series.Points.Any())
                    {
                        foreach (var point in series.Points)
                        {
                            PointDefinition pointDefinition = new(seriesDefinition)
                            {
                                Category = point.Category,
                                Value = point.Value,
                                X = point.X,
                                Y = point.Y,
                            };

                            seriesDefinition.Points.Add(pointDefinition);
                        }
                    }

                    Series.Add(seriesDefinition);
                }
            }

            return base.Analyse(chartNode, assets);
        }
    }
    public class ChartDefinition : ITitleOwner, IMarkerOwner, IValueLabelsContainer
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
        public List<SeriesDefinition> Series { get; set; } = new();
        public AxisDefinition CategoryAxis { get; set; }   // Not used for Pie/Doughnut
        public AxisDefinition ValueAxis { get; set; }   // Not used for Pie/Doughnut
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
        public bool ShowCategoryAxis { get; set; } //default true
        public bool ShowValueAxis { get; set; } //default true
        public int? FirstSliceAngle { get; set; } // Pie/Doughnut: 0..360
        public int? DoughnutHoleSize { get; set; } // Doughnut: 10..90
        public int? BubbleScale { get; set; } // Bubble: 0..300 (yüzde)
        public bool? Bubble3D { get; set; } // Bubble: true/false

        internal virtual bool Analyse(XmlNode chartNode, ExcelDocumentAssets assets)
        {
            if (!analyseChart(chartNode))
                return false;

            foreach (XmlNode node in chartNode.ChildNodes)
            {
                if (node.NodeType != XmlNodeType.Element)
                    continue;

                switch (node.Name.ToLowerInvariant())
                {
                    case var n when n == ExcelDocument.Tag.category_axis:
                        CategoryAxis = new();
                        analyseAxis(node, CategoryAxis);
                        break;
                    case var n when n == ExcelDocument.Tag.value_axis:
                        ValueAxis = new();
                        analyseAxis(node, ValueAxis);
                        break;
                    case var n when n == ExcelDocument.Tag.legend:
                        analyseLegend(node);
                        break;
                    case var n when n == ExcelDocument.Tag.grid:
                        analyseGrid(node);
                        break;
                    case var n when n == ExcelDocument.Tag.value_labels:
                        ValueLabels = AnalyseValueLabels(node);
                        break;
                    case var n when n == ExcelDocument.Tag.series:
                        analyseSeries(node);
                        break;
                    case var n when n == ExcelDocument.Tag.layout:
                        analyseLayout(node);
                        break;
                    case var n when n == ExcelDocument.Tag.view_3d:
                        analyse3dView(node);
                        break;
                    case var n when n == ExcelDocument.Tag.plot_area:
                        analysePlotArea(node);
                        break;
                    case var n when n == ExcelDocument.Tag.format:
                        ChartAreaFormat = AnalyseFormat(node);
                        break;
                    case var n when n == ExcelDocument.Tag.text_format:
                        TextFormat = AnalyseTextFormat(node);
                        break;
                    case var n when n == ExcelDocument.Tag.marker:
                        Marker = AnalyseMarker(node);
                        break;
                    case var n when n == ExcelDocument.Tag.title:
                        Title = AnalyseTitle(node);
                        break;
                    case var n when n == ExcelDocument.Tag.series_format:
                        SeriesDefaultFormat = AnalyseFormat(node);
                        break;
                    case var n when n == ExcelDocument.Tag.data_table:
                        DataTable = analyseDataTable(node);
                        break;
                    default:
                        break;
                }
            }

            this.ApplyTitleShortcut(chartNode);
            this.ApplyMarkerShortcut(chartNode);

            return true;
        }
        public ChartDefinition GetChartDefinition()
        {
            return this;
        }
        private bool analyseChart(XmlNode chartNode)
        {
            ChartType? chartType = ExcelXMLAttributeGetter.AsEnum<ChartType>(chartNode, ExcelDocument.Attribute.type);

            if (chartType is null)
            {
                return false;
            }

            Type = chartType.Value;
            ScatterStyle = ExcelXMLAttributeGetter.AsEnum<ScatterStyle>(chartNode, ExcelDocument.Attribute.scatter_style) ?? ScatterStyle.Marker;
            RadarStyle = ExcelXMLAttributeGetter.AsEnum<RadarStyle>(chartNode, ExcelDocument.Attribute.radar_style) ?? RadarStyle.Standard;
            Overlap = chartNode.ParseInt32WithMinMax(ExcelDocument.Attribute.overlap, -100, 100);
            GapWidth = chartNode.ParseInt32WithMinMax(ExcelDocument.Attribute.gap_width, 0, 500);
            Width = Dimension.Parse(ExcelXMLAttributeGetter.AsStringSafe(chartNode, ExcelDocument.Attribute.width)) ?? new Dimension(600, UnitType.Px);
            Height = Dimension.Parse(ExcelXMLAttributeGetter.AsStringSafe(chartNode, ExcelDocument.Attribute.height)) ?? new Dimension(400, UnitType.Px);
            BlanksDisplayedAs = ExcelXMLAttributeGetter.AsEnum<BlanksDisplayedAs>(chartNode, ExcelDocument.Attribute.display_blanks_as);
            GroupingType = ExcelXMLAttributeGetter.AsEnum<GroupingType>(chartNode, ExcelDocument.Attribute.grouping_type);
            AutoTitle = chartNode.ParseBool(ExcelDocument.Attribute.auto_title);
            PlotVisibleOnly = chartNode.ParseBool(ExcelDocument.Attribute.plot_visible_only);
            RoundedCorners = chartNode.ParseBool(ExcelDocument.Attribute.rounded_corners);
            VaryColors = chartNode.ParseBool(ExcelDocument.Attribute.vary_colors);
            ShowDataLabelsOverMaximum = chartNode.ParseBool(ExcelDocument.Attribute.show_data_labels_over_maximum);
            ShowCategoryAxis = chartNode.ParseBool(ExcelDocument.Attribute.show_category_axis) ?? true;
            ShowValueAxis = chartNode.ParseBool(ExcelDocument.Attribute.show_value_axis) ?? true;
            FirstSliceAngle = chartNode.ParseInt32WithMinMax(ExcelDocument.Attribute.first_slice_angle, 0, 360);
            DoughnutHoleSize = chartNode.ParseInt32WithMinMax(ExcelDocument.Attribute.doughnut_hole_size, 10, 90);
            BubbleScale = chartNode.ParseInt32WithMinMax(ExcelDocument.Attribute.bubble_scale, 0, 300);
            Bubble3D = chartNode.ParseBool(ExcelDocument.Attribute.bubble_3d);

            return true;
        }
        private void analysePlotArea(XmlNode node)
        {
            if (node is null)
                return;

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.NodeType != XmlNodeType.Element)
                    continue;

                switch (childNode.Name.ToLowerInvariant())
                {
                    case var n when n == ExcelDocument.Tag.format:
                        PlotAreaFormat = AnalyseFormat(childNode);
                        break;
                }
            }
        }
        private void analyseLegend(XmlNode node)
        {
            Legend = new()
            {
                Show = ExcelXMLAttributeGetter.AsBoolSafe(node, ExcelDocument.Attribute.show),
                Position = ExcelXMLAttributeGetter.AsEnum<LegendPosition>(node, ExcelDocument.Attribute.position) ?? LegendPosition.Right,
            };

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.NodeType != XmlNodeType.Element)
                    continue;

                switch (childNode.Name.ToLowerInvariant())
                {
                    case var n when n == ExcelDocument.Tag.format:
                        Legend.BoxFormat = AnalyseFormat(childNode);
                        break;
                    case var n when n == ExcelDocument.Tag.text_format:
                        Legend.LegendEntryTextFormat = AnalyseTextFormat(childNode);
                        break;
                }
            }
        }
        private void analyseGrid(XmlNode node)
        {
            Grid = new()
            {
                ShowHorizontal = ExcelXMLAttributeGetter.AsBoolSafe(node, ExcelDocument.Attribute.show_horizontal),
                ShowVertical = ExcelXMLAttributeGetter.AsBoolSafe(node, ExcelDocument.Attribute.show_vertical),
            };

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.NodeType != XmlNodeType.Element)
                    continue;

                switch (childNode.Name.ToLowerInvariant())
                {
                    case var n when n == ExcelDocument.Tag.horizontal_format:
                        Grid.HorizontalFormat = AnalyseFormat(childNode);
                        break;
                    case var n when n == ExcelDocument.Tag.vertical_format:
                        Grid.VerticalFormat = AnalyseFormat(childNode);
                        break;
                }
            }
        }
        internal static ValueLabelDefinition AnalyseValueLabels(XmlNode node)
        {
            ValueLabelDefinition valueLabelDefinition = new()
            {
                Position = ExcelXMLAttributeGetter.AsEnum<DataLabelPosition>(node, ExcelDocument.Attribute.position),
                Show = ChartDefinitionHelper.ParseBool(node, ExcelDocument.Attribute.show) ?? true,
                ShowLegendKey = ExcelXMLAttributeGetter.AsBoolSafe(node, ExcelDocument.Attribute.show_legend_key),
                ShowCategoryName = ExcelXMLAttributeGetter.AsBoolSafe(node, ExcelDocument.Attribute.show_category_name),
                ShowSeriesName = ExcelXMLAttributeGetter.AsBoolSafe(node, ExcelDocument.Attribute.show_series_name),
                ShowPercent = ExcelXMLAttributeGetter.AsBoolSafe(node, ExcelDocument.Attribute.show_percent),
                ShowBubbleSize = ExcelXMLAttributeGetter.AsBoolSafe(node, ExcelDocument.Attribute.show_bubble_size),
            };

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.NodeType != XmlNodeType.Element)
                    continue;

                switch (childNode.Name.ToLowerInvariant())
                {
                    case var n when n == ExcelDocument.Tag.text_format:
                        valueLabelDefinition.TextFormat = AnalyseTextFormat(childNode);
                        break;
                }
            }

            return valueLabelDefinition;
        }

        internal static class ColorUtil
        {
            /// <summary>
            /// “#RRGGBB” döndürmeye çalışır; değilse gelen değeri olduğu gibi bırakır.
            /// “rgb(r,g,b)” formatını da destekler.
            /// </summary>
            public static string ToHex(string color)
            {
                if (string.IsNullOrWhiteSpace(color)) return null;

                // #RRGGBB ise aynen bırak
                if (System.Text.RegularExpressions.Regex.IsMatch(color, "^#([0-9a-fA-F]{6})$"))
                    return color;

                // rgb(255,0,0) -> #FF0000
                var m = System.Text.RegularExpressions.Regex.Match(
                    color, @"rgb\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                );
                if (m.Success)
                {
                    int clamp(int v) => Math.Max(0, Math.Min(255, v));
                    int r = clamp(int.Parse(m.Groups[1].Value));
                    int g = clamp(int.Parse(m.Groups[2].Value));
                    int b = clamp(int.Parse(m.Groups[3].Value));
                    return $"#{r:X2}{g:X2}{b:X2}";
                }

                // Başka bir isimlendirme ise şimdilik olduğu gibi dön
                return color;
            }
        }


        private void analyseLayout(XmlNode node)
        {
            ChartLayoutDefinition layout = new()
            {
                X = node.ParseInt32WithMinMax(ExcelDocument.Attribute.x, 0, 100),
                Y = node.ParseInt32WithMinMax(ExcelDocument.Attribute.y, 0, 100),
                Width = node.ParseInt32WithMinMax(ExcelDocument.Attribute.width, 0, 100),
                Height = node.ParseInt32WithMinMax(ExcelDocument.Attribute.height, 0, 100),
                XMode = ExcelXMLAttributeGetter.AsEnum<LayoutMode>(node, ExcelDocument.Attribute.x_mode),
                YMode = ExcelXMLAttributeGetter.AsEnum<LayoutMode>(node, ExcelDocument.Attribute.y_mode),
                Target = ExcelXMLAttributeGetter.AsEnum<LayoutTarget>(node, ExcelDocument.Attribute.target),
            };

            if (layout.HasManualLayout)
                Layout = layout;
        }
        private void analyse3dView(XmlNode node)
        {
            ThreeDViewDefinition threeDView = new()
            {
                Shape = ExcelXMLAttributeGetter.AsEnum<ThreeDShape>(node, ExcelDocument.Attribute.shape),
                RotationX = node.ParseInt32WithMinMax(ExcelDocument.Attribute.rotation_x),
                RotationY = node.ParseInt32WithMinMax(ExcelDocument.Attribute.rotation_y),
                Perspective = node.ParseInt32WithMinMax(ExcelDocument.Attribute.perspective),
                DepthPercent = node.ParseInt32WithMinMax(ExcelDocument.Attribute.depth_percent),
                HeightPercent = node.ParseInt32WithMinMax(ExcelDocument.Attribute.height_percent),
                GapWidth = node.ParseInt32WithMinMax(ExcelDocument.Attribute.gap_width),
                GapDepth = node.ParseInt32WithMinMax(ExcelDocument.Attribute.gap_depth),
                RightAngleAxes = node.ParseBool(ExcelDocument.Attribute.right_angle_axes),
                ShowFloor = node.ParseBool(ExcelDocument.Attribute.show_floor),
                ShowBackWall = node.ParseBool(ExcelDocument.Attribute.show_back_wall),
                ShowSideWall = node.ParseBool(ExcelDocument.Attribute.show_side_wall),
            };

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.NodeType != XmlNodeType.Element)
                    continue;

                switch (childNode.Name.ToLowerInvariant())
                {
                    case var n when n == ExcelDocument.Tag.format:
                        threeDView.DefaultFormat = AnalyseFormat(childNode);
                        break;
                    case var n when n == ExcelDocument.Tag.floor_format:
                        threeDView.FloorFormat = AnalyseFormat(childNode);
                        break;
                    case var n when n == ExcelDocument.Tag.back_wall_format:
                        threeDView.BackWallFormat = AnalyseFormat(childNode);
                        break;
                    case var n when n == ExcelDocument.Tag.side_wall_format:
                        threeDView.SideWallFormat = AnalyseFormat(childNode);
                        break;
                }
            }

            if (threeDView.DefaultFormat?.FillDefinition is null)
            {
                if ((threeDView.ShowFloor is true && threeDView.FloorFormat?.FillDefinition is null) ||
                    (threeDView.ShowSideWall is true && threeDView.SideWallFormat?.FillDefinition is null) ||
                    (threeDView.ShowBackWall is true && threeDView.BackWallFormat?.FillDefinition is null))
                {
                    threeDView.DefaultFormat ??= new();
                    threeDView.DefaultFormat.FillDefinition = new()
                    {
                        SolidFillDefinition = new()
                    };
                }
            }

            ThreeDView = threeDView;
        }
        private void analyseSeries(XmlNode node)
        {
            SeriesDefinition series = new(this);

            Series.Add(series);

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.NodeType != XmlNodeType.Element)
                    continue;

                switch (childNode.Name.ToLowerInvariant())
                {
                    case var n when n == ExcelDocument.Tag.point:
                        {
                            PointDefinition point = null;

                            switch (Type)
                            {
                                case ChartType.Bar:
                                case ChartType.Column:
                                case ChartType.Line:
                                case ChartType.Area:
                                case ChartType.Pie:
                                case ChartType.Doughnut:
                                case ChartType.Radar:
                                    point = parseCategoricalPoint(childNode, series);
                                    break;
                                case ChartType.Scatter:
                                    point = parseScatterPoint(childNode, series);
                                    break;
                                case ChartType.Bubble:
                                    point = parseBubblePoint(childNode, series);
                                    break;
                                default:
                                    break;
                            }

                            if (point is not null)
                                series.Points.Add(point);
                        }
                        break;
                    case var n when n == ExcelDocument.Tag.format:
                        series.Format = AnalyseFormat(childNode);
                        break;
                    case var n when n == ExcelDocument.Tag.marker:
                        series.Marker = AnalyseMarker(childNode);
                        break;
                    case var n when n == ExcelDocument.Tag.title:
                        series.Title = AnalyseTitle(childNode);
                        break;
                    case var n when n == ExcelDocument.Tag.value_labels:
                        series.ValueLabels = AnalyseValueLabels(childNode);
                        break;
                }
            }

            string color = ExcelXMLAttributeGetter.AsStringSafe(node, ExcelDocument.Attribute.color);

            if (!string.IsNullOrWhiteSpace(color))
            {
                int? colorTransparency = node.ParseInt32WithMinMax(ExcelDocument.Attribute.color_transparency, 0, 100);

                if (colorTransparency is null && ((Type == ChartType.Radar && RadarStyle == RadarStyle.Filled)))
                    colorTransparency = 60;

                if (series.Format?.FillDefinition is null)
                {
                    series.Format ??= new();
                    series.Format.FillDefinition = new()
                    {
                        SolidFillDefinition = new()
                        {
                            Color = new ColorA(color, colorTransparency)
                        }
                    };
                }

                if (IsLineBased && series.Format?.LineDefinition is null)
                {
                    series.Format ??= new();
                    series.Format.LineDefinition = new()
                    {
                        Visible = true,
                        DashPreset = DashPreset.Solid,
                        SolidLineDefinition = new()
                        {
                            Color = new ColorA(color, colorTransparency)
                        }
                    };
                }
            }

            series.ApplyTitleShortcut(node);
            series.ApplyMarkerShortcut(node);
        }
        private static void analyseAxis(XmlNode node, AxisDefinition axis)
        {
            if (axis is null)
                return;

            axis.Min = node.ParseDoubleWithMinMax(ExcelDocument.Attribute.min);

            if (ExcelXMLAttributeGetter.AsDouble(node, ExcelDocument.Attribute.max, out double max) && (!axis.Min.HasValue || axis.Min.Value < max))
                axis.Max = max;

            if (ExcelXMLAttributeGetter.AsDouble(node, ExcelDocument.Attribute.major_unit, out double major_unit) && major_unit > 0)
                axis.MajorUnit = major_unit;

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.NodeType != XmlNodeType.Element)
                    continue;

                switch (childNode.Name.ToLowerInvariant())
                {
                    case var n when n == ExcelDocument.Tag.format:
                        axis.AxisLineFormat = AnalyseFormat(childNode);
                        break;
                    case var n when n == ExcelDocument.Tag.text_format:
                        axis.TickLabelTextFormat = AnalyseTextFormat(childNode);
                        break;
                    case var n when n == ExcelDocument.Tag.title:
                        axis.Title = AnalyseTitle(childNode);
                        break;
                }
            }

            axis.ApplyTitleShortcut(node);
        }
        private static DataTableDefinition analyseDataTable(XmlNode node)
        {
            DataTableDefinition dataTable = new()
            {
                Show = ChartDefinitionHelper.ParseBool(node, ExcelDocument.Attribute.show) ?? true,
                ShowHorizontalBorder = ChartDefinitionHelper.ParseBool(node, ExcelDocument.Attribute.show_horizontal_border) ?? true,
                ShowVerticalBorder = ChartDefinitionHelper.ParseBool(node, ExcelDocument.Attribute.show_vertical_border) ?? true,
                ShowOutlineBorder = ChartDefinitionHelper.ParseBool(node, ExcelDocument.Attribute.show_outline_border) ?? true,
                ShowLegendKey = ChartDefinitionHelper.ParseBool(node, ExcelDocument.Attribute.show_legend_key) ?? true,
            };

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.NodeType != XmlNodeType.Element)
                    continue;

                switch (childNode.Name.ToLowerInvariant())
                {
                    case var n when n == ExcelDocument.Tag.format:
                        dataTable.BoxFormat = AnalyseFormat(childNode);
                        break;
                    case var n when n == ExcelDocument.Tag.text_format:
                        dataTable.TextFormat = AnalyseTextFormat(childNode);
                        break;
                }
            }

            return dataTable;
        }
        internal static TitleDefinition AnalyseTitle(XmlNode node)
        {
            TitleDefinition titleDefinition = new()
            {
                Text = ExcelXMLAttributeGetter.AsString(node, ExcelDocument.Attribute.text),
            };

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.NodeType != XmlNodeType.Element)
                    continue;

                switch (childNode.Name.ToLowerInvariant())
                {
                    case var n when n == ExcelDocument.Tag.text_format:
                        titleDefinition.TextFormat = AnalyseTextFormat(childNode);
                        break;
                }
            }

            return titleDefinition;
        }
        internal static FormatDefinition AnalyseFormat(XmlNode node)
        {
            if (node is null)
                return null;

            FormatDefinition formatDefinition = new();

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.NodeType != XmlNodeType.Element)
                    continue;

                switch (childNode.Name.ToLowerInvariant())
                {
                    case var n when n == ExcelDocument.Tag.fill: formatDefinition.FillDefinition = analyseFill(childNode); break;
                    case var n when n == ExcelDocument.Tag.line: formatDefinition.LineDefinition = analyseLine(childNode); break;
                    case var n when n == ExcelDocument.Tag.effects: formatDefinition.EffectsDefinition = analyseEffects(childNode); break;
                    default:
                        break;
                }
            }

            return formatDefinition;
        }
        internal static MarkerDefinition AnalyseMarker(XmlNode node)
        {
            MarkerDefinition markerDefinition = new()
            {
                Type = ExcelXMLAttributeGetter.AsEnum<MarkerType>(node, ExcelDocument.Attribute.type),
                Size = node.ParseInt32WithMinMax(ExcelDocument.Attribute.size, 2, 72)
            };

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.NodeType != XmlNodeType.Element)
                    continue;

                switch (childNode.Name.ToLowerInvariant())
                {
                    case var n when n == ExcelDocument.Tag.format:
                        markerDefinition.Format = AnalyseFormat(childNode);
                        break;
                }
            }

            return markerDefinition;
        }
        internal static TextFormatDefinition AnalyseTextFormat(XmlNode node)
        {
            if (node is null)
                return null;

            TextFormatDefinition textFormatDefinition = new()
            {
                AlignHorizontal = ExcelXMLAttributeGetter.AsEnum<HorizontalAlign>(node, ExcelDocument.Attribute.align_horizontal),
                AlignVertical = ExcelXMLAttributeGetter.AsEnum<VerticalAlign>(node, ExcelDocument.Attribute.align_vertical),
                Wrap = ExcelXMLAttributeGetter.AsEnum<TextWrapPreset>(node, ExcelDocument.Attribute.wrap),
                MarginTop = Dimension.Parse(ExcelXMLAttributeGetter.AsStringSafe(node, ExcelDocument.Attribute.margin_top)),
                MarginRight = Dimension.Parse(ExcelXMLAttributeGetter.AsStringSafe(node, ExcelDocument.Attribute.margin_right)),
                MarginBottom = Dimension.Parse(ExcelXMLAttributeGetter.AsStringSafe(node, ExcelDocument.Attribute.margin_bottom)),
                MarginLeft = Dimension.Parse(ExcelXMLAttributeGetter.AsStringSafe(node, ExcelDocument.Attribute.margin_left)),
                Rotate = node.ParseInt32WithMinMax(ExcelDocument.Attribute.rotate, -90, 90),
            };

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.NodeType != XmlNodeType.Element)
                    continue;

                switch (childNode.Name.ToLowerInvariant())
                {
                    case var n when n == ExcelDocument.Tag.fill: textFormatDefinition.TextFillDefinition = analyseFill(childNode); break;
                    case var n when n == ExcelDocument.Tag.line: textFormatDefinition.TextOutlineDefinition = analyseLine(childNode); break;
                    case var n when n == ExcelDocument.Tag.effects: textFormatDefinition.EffectsDefinition = analyseEffects(childNode); break;
                    case var n when n == ExcelDocument.Tag.font: textFormatDefinition.FontDefinition = analyseFont(childNode); break;
                    default:
                        break;
                }
            }

            return textFormatDefinition;
        }
        private static FillDefinition analyseFill(XmlNode node)
        {
            if (node is null)
                return null;

            FillDefinition fillDefinition = new();

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.NodeType != XmlNodeType.Element)
                    continue;

                switch (childNode.Name.ToLowerInvariant())
                {
                    case var n when n == ExcelDocument.Tag.solid: fillDefinition.SolidFillDefinition = analyseSolid(childNode); break;
                    case var n when n == ExcelDocument.Tag.pattern: fillDefinition.PatternFillDefinition = analysePattern(childNode); break;
                    case var n when n == ExcelDocument.Tag.gradient: fillDefinition.GradientFillDefinition = analyseGradient(childNode); break;
                    default:
                        break;
                }
            }

            return fillDefinition;
        }
        private static SolidDefinition analyseSolid(XmlNode node)
        {
            if (node is null)
                return null;

            ColorA color = ColorA.Parse(node, ExcelDocument.Attribute.color, ExcelDocument.Attribute.transparency);

            if (color is null)
                return null;

            return new SolidDefinition() { Color = color };
        }
        private static PatternDefinition analysePattern(XmlNode node)
        {
            if (node is null)
                return null;

            PatternDefinition patternDefinition = new()
            {
                Preset = ExcelXMLAttributeGetter.AsEnum<PatternFillPreset>(node, ExcelDocument.Attribute.preset) ?? PatternFillPreset.Cross,
                ForegroundColor = ColorA.Parse(node, ExcelDocument.Attribute.foreground_color, ExcelDocument.Attribute.foreground_transparency),
                BackgroundColor = ColorA.Parse(node, ExcelDocument.Attribute.background_color, ExcelDocument.Attribute.background_transparency),
            };

            return patternDefinition;
        }
        private static GradientDefinition analyseGradient(XmlNode node)
        {
            if (node is null)
                return null;

            GradientDefinition gradientDefinition = new()
            {
                Angle = node.ParseInt32WithMinMax(ExcelDocument.Attribute.angle, 0, 360),
                Scaled = node.ParseBool(ExcelDocument.Attribute.scaled),
            };

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.NodeType != XmlNodeType.Element)
                    continue;

                switch (childNode.Name.ToLowerInvariant())
                {
                    case var n when n == ExcelDocument.Tag.stop:
                        GradientStopDefinition stop = analyseGradientStop(childNode);

                        if (stop is not null)
                            gradientDefinition.Stops.Add(stop);
                        break;
                    default:
                        break;
                }
            }

            if (gradientDefinition.Stops.Any())
            {
                gradientDefinition.Stops =
                    gradientDefinition.Stops
                        .GroupBy(s => Math.Max(0, Math.Min(100, s.Position)))
                        .Select(g => g.Last())
                        .OrderBy(s => s.Position)
                        .ToList();
            }

            return gradientDefinition;
        }
        private static GradientStopDefinition analyseGradientStop(XmlNode node)
        {
            if (node is null)
                return null;

            int? position = node.ParseInt32WithMinMax(ExcelDocument.Attribute.position, 0, 100);

            if (position is null)
                return null;

            return new GradientStopDefinition()
            {
                Position = position.Value,
                Color = ColorA.Parse(node, ExcelDocument.Attribute.color, ExcelDocument.Attribute.transparency)
            };
        }
        private static LineDefinition analyseLine(XmlNode node)
        {
            if (node is null)
                return null;

            LineDefinition lineDefinition = new()
            {
                Visible = node.ParseBool(ExcelDocument.Attribute.visible) ?? true,
                Width = Dimension.Parse(ExcelXMLAttributeGetter.AsStringSafe(node, ExcelDocument.Attribute.width)),
                DashPreset = ExcelXMLAttributeGetter.AsEnum<DashPreset>(node, ExcelDocument.Attribute.dash),
                CompoundPreset = ExcelXMLAttributeGetter.AsEnum<CompoundLinePreset>(node, ExcelDocument.Attribute.compound),
                CapPreset = ExcelXMLAttributeGetter.AsEnum<LineCapPreset>(node, ExcelDocument.Attribute.cap),
                JoinPreset = ExcelXMLAttributeGetter.AsEnum<LineJoinPreset>(node, ExcelDocument.Attribute.join),
                Begin_Arrow_Type = ExcelXMLAttributeGetter.AsEnum<LineEndPreset>(node, ExcelDocument.Attribute.begin_arrow_type),
                Begin_Arrow_Width = ExcelXMLAttributeGetter.AsEnum<LineEndWidthPreset>(node, ExcelDocument.Attribute.begin_arrow_width),
                Begin_Arrow_Length = ExcelXMLAttributeGetter.AsEnum<LineEndLengthPreset>(node, ExcelDocument.Attribute.begin_arrow_length),
                End_Arrow_Type = ExcelXMLAttributeGetter.AsEnum<LineEndPreset>(node, ExcelDocument.Attribute.end_arrow_type),
                End_Arrow_Width = ExcelXMLAttributeGetter.AsEnum<LineEndWidthPreset>(node, ExcelDocument.Attribute.end_arrow_width),
                End_Arrow_Length = ExcelXMLAttributeGetter.AsEnum<LineEndLengthPreset>(node, ExcelDocument.Attribute.end_arrow_length),
                JoinMiterLimit = node.ParseInt32WithMinMax(ExcelDocument.Attribute.join_miter_limit, 1, 500),
            };

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.NodeType != XmlNodeType.Element)
                    continue;

                switch (childNode.Name.ToLowerInvariant())
                {
                    case var n when n == ExcelDocument.Tag.solid: lineDefinition.SolidLineDefinition = analyseSolid(childNode); break;
                    case var n when n == ExcelDocument.Tag.gradient: lineDefinition.GradientLineDefinition = analyseGradient(childNode); break;
                    default:
                        break;
                }
            }

            if (lineDefinition.SolidLineDefinition is null && lineDefinition.GradientLineDefinition is null)
            {
                lineDefinition.SolidLineDefinition = new SolidDefinition() { Color = new("#000000", null) };
            }

            return lineDefinition;
        }
        private static EffectsDefinition analyseEffects(XmlNode node)
        {
            if (node is null)
                return null;

            EffectsDefinition effectsDefinition = new();

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.NodeType != XmlNodeType.Element)
                    continue;

                switch (childNode.Name.ToLowerInvariant())
                {
                    case var n when n == ExcelDocument.Tag.shadow: effectsDefinition.ShadowDefinition = analyseShadow(childNode); break;
                    case var n when n == ExcelDocument.Tag.glow: effectsDefinition.GlowDefinition = analyseGlow(childNode); break;
                    case var n when n == ExcelDocument.Tag.soft_edges: effectsDefinition.SoftEdgesDefinition = analyseSoftEdges(childNode); break;
                    case var n when n == ExcelDocument.Tag.format_3D: effectsDefinition.Format3dDefinition = analyseFormat3D(childNode); break;
                    case var n when n == ExcelDocument.Tag.reflection: effectsDefinition.ReflectionDefinition = analyseReflection(childNode); break;
                    default:
                        break;
                }
            }

            return effectsDefinition;
        }
        private static ShadowDefinition analyseShadow(XmlNode node)
        {
            if (node is null)
                return null;

            ShadowDefinition shadowDefinition = new()
            {
                Type = ExcelXMLAttributeGetter.AsEnum<ShadowType>(node, ExcelDocument.Attribute.type),
                Preset = ExcelXMLAttributeGetter.AsEnum<ShadowPreset>(node, ExcelDocument.Attribute.preset),
                Color = ColorA.Parse(node, ExcelDocument.Attribute.color, ExcelDocument.Attribute.transparency),
                BlurRadius = Dimension.Parse(ExcelXMLAttributeGetter.AsStringSafe(node, ExcelDocument.Attribute.blur_radius)),
                Distance = Dimension.Parse(ExcelXMLAttributeGetter.AsStringSafe(node, ExcelDocument.Attribute.distance)),
                Angle = node.ParseInt32WithMinMax(ExcelDocument.Attribute.angle, 0, 360),
            };

            return shadowDefinition;
        }
        private static GlowDefinition analyseGlow(XmlNode node)
        {
            if (node is null)
                return null;

            GlowDefinition glowDefinition = new()
            {
                Color = ColorA.Parse(node, ExcelDocument.Attribute.color, ExcelDocument.Attribute.transparency),
                Size = Dimension.Parse(ExcelXMLAttributeGetter.AsStringSafe(node, ExcelDocument.Attribute.size)),
            };

            return glowDefinition;
        }
        private static SoftEdgesDefinition analyseSoftEdges(XmlNode node)
        {
            if (node is null)
                return null;

            SoftEdgesDefinition softEdgesDefinition = new()
            {
                Size = Dimension.Parse(ExcelXMLAttributeGetter.AsStringSafe(node, ExcelDocument.Attribute.size)),
            };

            return softEdgesDefinition;
        }
        private static ReflectionDefinition analyseReflection(XmlNode node)
        {
            if (node is null)
                return null;

            ReflectionDefinition reflectionDefinition = new()
            {
                Blur = Dimension.Parse(ExcelXMLAttributeGetter.AsStringSafe(node, ExcelDocument.Attribute.blur_radius)),
                Distance = Dimension.Parse(ExcelXMLAttributeGetter.AsStringSafe(node, ExcelDocument.Attribute.distance)),
                StartTransparency = node.ParseInt32WithMinMax(ExcelDocument.Attribute.start_transparency, 0, 100),
                EndTransparency = node.ParseInt32WithMinMax(ExcelDocument.Attribute.end_transparency, 0, 100),
                StartPosition = node.ParseInt32WithMinMax(ExcelDocument.Attribute.start_position, 0, 100),
                EndPosition = node.ParseInt32WithMinMax(ExcelDocument.Attribute.end_position, 0, 100),
            };

            if (reflectionDefinition.StartPosition is not null && reflectionDefinition.EndPosition is not null && reflectionDefinition.StartPosition >= reflectionDefinition.EndPosition)
                reflectionDefinition.StartPosition = 0;

            return reflectionDefinition;
        }
        private static Format3DDefinition analyseFormat3D(XmlNode node)
        {
            if (node is null)
                return null;

            Format3DDefinition format3DDefinition = new()
            {
                Material = ExcelXMLAttributeGetter.AsEnum<MaterialPreset>(node, ExcelDocument.Attribute.material),
            };

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.NodeType != XmlNodeType.Element)
                    continue;

                switch (childNode.Name.ToLowerInvariant())
                {
                    case var n when n == ExcelDocument.Tag.bevel: format3DDefinition.BevelDefinition = analyseBevel(childNode); break;
                    case var n when n == ExcelDocument.Tag.lighting: format3DDefinition.LightingDefinition = analyseLighting(childNode); break;
                    default:
                        break;
                }
            }

            if (format3DDefinition.BevelDefinition?.TopPreset is null && format3DDefinition.BevelDefinition?.BottomPreset is null)
            {
                format3DDefinition.BevelDefinition ??= new();

                if (format3DDefinition.BevelDefinition.TopHeight is null)
                    format3DDefinition.BevelDefinition.TopHeight = new Dimension(6, UnitType.Pt);

                if (format3DDefinition.BevelDefinition.TopWidth is null)
                    format3DDefinition.BevelDefinition.TopWidth = new Dimension(6, UnitType.Pt);

                if (format3DDefinition.BevelDefinition.TopPreset is null)
                    format3DDefinition.BevelDefinition.TopPreset = BevelPreset.Circle;
            }

            return format3DDefinition;
        }
        private static BevelDefinition analyseBevel(XmlNode node)
        {
            if (node is null)
                return null;

            BevelDefinition bevelDefinition = new()
            {
                TopWidth = Dimension.Parse(ExcelXMLAttributeGetter.AsString(node, ExcelDocument.Attribute.top_width)),
                TopHeight = Dimension.Parse(ExcelXMLAttributeGetter.AsString(node, ExcelDocument.Attribute.top_height)),
                BottomWidth = Dimension.Parse(ExcelXMLAttributeGetter.AsString(node, ExcelDocument.Attribute.bottom_width)),
                BottomHeight = Dimension.Parse(ExcelXMLAttributeGetter.AsString(node, ExcelDocument.Attribute.bottom_height)),
                TopPreset = ExcelXMLAttributeGetter.AsEnum<BevelPreset>(node, ExcelDocument.Attribute.top_preset),
                BottomPreset = ExcelXMLAttributeGetter.AsEnum<BevelPreset>(node, ExcelDocument.Attribute.bottom_preset),
            };

            return bevelDefinition;
        }
        private static LightingDefinition analyseLighting(XmlNode node)
        {
            if (node is null)
                return null;

            LightingDefinition lightingDefinition = new()
            {
                Angle = node.ParseInt32WithMinMax(ExcelDocument.Attribute.angle, 0, 360),
                Preset = ExcelXMLAttributeGetter.AsEnum<LightingPreset>(node, ExcelDocument.Attribute.preset),
                Direction = ExcelXMLAttributeGetter.AsEnum<LightingDirection>(node, ExcelDocument.Attribute.direction),
            };

            return lightingDefinition;
        }
        private static FontDefinition analyseFont(XmlNode node)
        {
            if (node is null)
                return null;

            FontDefinition fontDefinition = new()
            {
                Family = ExcelXMLAttributeGetter.AsString(node, ExcelDocument.Attribute.family),
                Size = Dimension.Parse(ExcelXMLAttributeGetter.AsString(node, ExcelDocument.Attribute.size)),
                Bold = node.ParseBool(ExcelDocument.Attribute.bold),
                Italic = node.ParseBool(ExcelDocument.Attribute.italic),
                Underline = ExcelXMLAttributeGetter.AsEnum<TextUnderlineStyle>(node, ExcelDocument.Attribute.underline),
                Strike = ExcelXMLAttributeGetter.AsEnum<TextStrikeStyle>(node, ExcelDocument.Attribute.strike),
                Caps = ExcelXMLAttributeGetter.AsEnum<TextCapsStyle>(node, ExcelDocument.Attribute.caps),
                Kerning = Dimension.Parse(ExcelXMLAttributeGetter.AsString(node, ExcelDocument.Attribute.kerning)),
                Spacing = Dimension.Parse(ExcelXMLAttributeGetter.AsString(node, ExcelDocument.Attribute.spacing)),
            };

            return fontDefinition;
        }
        private static PointDefinition parseCategoricalPoint(XmlNode node, SeriesDefinition seriesDefinition)
        {
            string category = ExcelXMLAttributeGetter.AsString(node, ExcelDocument.Attribute.category);

            if (string.IsNullOrWhiteSpace(category))
                return null;

            if (!ExcelXMLAttributeGetter.AsDouble(node, ExcelDocument.Attribute.value, out double value))
                return null;

            return new PointDefinition(node, seriesDefinition)
            {
                Category = category,
                Value = value,
            };
        }
        private static PointDefinition parseScatterPoint(XmlNode node, SeriesDefinition seriesDefinition)
        {
            if (!ExcelXMLAttributeGetter.AsDouble(node, ExcelDocument.Attribute.x, out double x))
                return null;

            if (!ExcelXMLAttributeGetter.AsDouble(node, ExcelDocument.Attribute.y, out double y))
                return null;

            return new PointDefinition(node, seriesDefinition)
            {
                X = x,
                Y = y,
            };
        }
        private static PointDefinition parseBubblePoint(XmlNode node, SeriesDefinition seriesDefinition)
        {
            if (!ExcelXMLAttributeGetter.AsDouble(node, ExcelDocument.Attribute.x, out double x))
                return null;

            if (!ExcelXMLAttributeGetter.AsDouble(node, ExcelDocument.Attribute.y, out double y))
                return null;

            if (!ExcelXMLAttributeGetter.AsDouble(node, ExcelDocument.Attribute.size, out double size))
                return null;

            return new PointDefinition(node, seriesDefinition)
            {
                X = x,
                Y = y,
                Size = size,
            };
        }
    }
    public class SeriesDefinition : ITitleOwner, IMarkerOwner, IValueLabelsContainer
    {
        public TitleDefinition Title { get; set; }
        public ColorA ColorA { get { return Format?.FillDefinition?.SolidFillDefinition?.Color; } }
        public string Color { get { return ColorA?.Color; } }
        public int? ColorTransparency { get { return ColorA?.Transparency; } }
        public List<PointDefinition> Points { get; set; } = new();
        public FormatDefinition Format { get; set; }
        public MarkerDefinition Marker { get; set; }
        public ValueLabelDefinition ValueLabels { get; set; } //<value-labels>
        public ChartDefinition ChartDefinition { get; set; }
        public SeriesDefinition(ChartDefinition chartDefinition)
        {
            ChartDefinition = chartDefinition;
        }
        public ChartDefinition GetChartDefinition()
        {
            return ChartDefinition;
        }
    }
    public class PointDefinition : IValueLabelsContainer
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

        public  /*Ctor*/                PointDefinition(SeriesDefinition seriesDefinition)
        {
            SeriesDefinition = seriesDefinition;
        }
        public  /*Ctor*/                PointDefinition(XmlNode node, SeriesDefinition seriesDefinition) : this(seriesDefinition)
        {
            if (node is null)
                return;

            MarkerType = ExcelXMLAttributeGetter.AsEnum<MarkerType>(node, ExcelDocument.Attribute.marker_type);
            MarkerSize = node.ParseInt32WithMinMax(ExcelDocument.Attribute.marker_size, 2, 72);

            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.NodeType != XmlNodeType.Element)
                    continue;

                switch (childNode.Name.ToLowerInvariant())
                {
                    case var n when n == ExcelDocument.Tag.format:
                        PointFormat = ChartDefinition.AnalyseFormat(childNode);
                        break;
                    case var n when n == ExcelDocument.Tag.value_labels:
                        ValueLabels = ChartDefinition.AnalyseValueLabels(childNode);
                        break;
                }
            }
        }
        public ChartDefinition GetChartDefinition()
        {
            return SeriesDefinition?.GetChartDefinition();
        }
    }
    public class AxisDefinition : ITitleOwner
    {
        public TitleDefinition Title { get; set; } // <title>
        public FormatDefinition AxisLineFormat { get; set; } // <format>
        public TextFormatDefinition TickLabelTextFormat { get; set; } // <text-format>

        // Only valid for numeric axes
        public double? Min { get; set; }
        public double? Max { get; set; }
        public double? MajorUnit { get; set; }
    }
    public class LegendDefinition
    {
        public bool Show { get; set; }
        public LegendPosition Position { get; set; }
        public FormatDefinition BoxFormat { get; set; }
        public TextFormatDefinition LegendEntryTextFormat { get; set; }
    }
    public class GridDefinition
    {
        public bool ShowHorizontal { get; set; } = true;
        public bool ShowVertical { get; set; } = true;
        public FormatDefinition HorizontalFormat { get; set; }
        public FormatDefinition VerticalFormat { get; set; }
    }
    public class ValueLabelDefinition
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
    public class ChartLayoutDefinition
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
    public class ThreeDViewDefinition
    {
        public int? RotationX { get; set; } // -90..90
        public int? RotationY { get; set; } // 0..360
        public bool? RightAngleAxes { get; set; }
        public int? Perspective { get; set; } // 0..240
        public int? DepthPercent { get; set; } // 20..2000
        public int? HeightPercent { get; set; } // 5..500

        public int? GapWidth { get; set; } // 0..500
        public int? GapDepth { get; set; } // 0..500
        public ThreeDShape? Shape { get; set; }

        public bool? ShowFloor { get; set; }
        public bool? ShowBackWall { get; set; }
        public bool? ShowSideWall { get; set; }

        public FormatDefinition DefaultFormat { get; set; }
        public FormatDefinition FloorFormat { get; set; }
        public FormatDefinition BackWallFormat { get; set; }
        public FormatDefinition SideWallFormat { get; set; }
    }
    public class MarkerDefinition
    {
        public MarkerType? Type { get; set; }
        public int? Size { get; set; }
        public FormatDefinition Format { get; set; }
    }
    public class TitleDefinition
    {
        public string Text { get; set; }
        public TextFormatDefinition TextFormat { get; set; }

        public static implicit operator string(TitleDefinition d) => d?.Text;
    }
    public class DataTableDefinition
    {
        public bool Show { get; set; }
        public bool ShowHorizontalBorder { get; set; }
        public bool ShowVerticalBorder { get; set; }
        public bool ShowOutlineBorder { get; set; }
        public bool ShowLegendKey { get; set; }
        public FormatDefinition BoxFormat { get; set; }
        public TextFormatDefinition TextFormat { get; set; }
    }
    public class FormatDefinition
    {
        public FillDefinition FillDefinition { get; set; }
        public LineDefinition LineDefinition { get; set; }
        public EffectsDefinition EffectsDefinition { get; set; }
    }
    public class TextFormatDefinition
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
    public class FillDefinition
    {
        public SolidDefinition SolidFillDefinition { get; set; }
        public PatternDefinition PatternFillDefinition { get; set; }
        public GradientDefinition GradientFillDefinition { get; set; }
    }
    public class SolidDefinition
    {
        public ColorA Color { get; set; }
    }
    public class PatternDefinition
    {
        public PatternFillPreset Preset { get; set; }
        public ColorA ForegroundColor { get; set; }
        public ColorA BackgroundColor { get; set; }
    }
    public class GradientDefinition
    {
        public int? Angle { get; set; }
        public bool? Scaled { get; set; }
        public List<GradientStopDefinition> Stops { get; set; } = new();
    }
    public class GradientStopDefinition
    {
        public int Position { get; set; }
        public ColorA Color { get; set; }
    }
    public class LineDefinition
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
    public class EffectsDefinition
    {
        public ShadowDefinition ShadowDefinition { get; set; }
        public GlowDefinition GlowDefinition { get; set; }
        public SoftEdgesDefinition SoftEdgesDefinition { get; set; }
        public Format3DDefinition Format3dDefinition { get; set; }
        public ReflectionDefinition ReflectionDefinition { get; set; }
    }
    public class ShadowDefinition
    {
        public ShadowType? Type { get; set; }
        public ShadowPreset? Preset { get; set; }
        public ColorA Color { get; set; }
        public Dimension BlurRadius { get; set; }
        public Dimension Distance { get; set; }
        public int? Angle { get; set; }
    }
    public class GlowDefinition
    {
        public ColorA Color { get; set; }
        public Dimension Size { get; set; }
    }
    public class SoftEdgesDefinition
    {
        public Dimension Size { get; set; }
    }
    public class ReflectionDefinition
    {
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
    public class Format3DDefinition
    {
        public MaterialPreset? Material { get; set; }
        public BevelDefinition BevelDefinition { get; set; }
        public LightingDefinition LightingDefinition { get; set; }
    }
    public class BevelDefinition
    {
        public bool HasData { get { return TopPreset is not null || BottomPreset is not null; } }
        public Dimension TopWidth { get; set; }
        public Dimension TopHeight { get; set; }
        public Dimension BottomWidth { get; set; }
        public Dimension BottomHeight { get; set; }
        public BevelPreset? TopPreset { get; set; }
        public BevelPreset? BottomPreset { get; set; }
    }
    public class LightingDefinition
    {
        public LightingPreset? Preset { get; set; }
        public LightingDirection? Direction { get; set; }
        public int? Angle { get; set; }
    }
    public class FontDefinition
    {
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
    public class Dimension
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
    public class ColorA
    {
        public string Color { get; set; }
        public int? Transparency { get; set; }

        public  /*Ctor*/    ColorA(string color, int? transparency)
        {
            Color = color;
            Transparency = transparency;
        }

        internal static ColorA Parse(XmlNode node, ExcelDocument.Attribute colorAttr, ExcelDocument.Attribute transparencyAttr)
        {
            if (node is null)
                return null;

            string color = ExcelXMLAttributeGetter.AsString(node, colorAttr);

            if (string.IsNullOrWhiteSpace(color))
                return null;

            return new ColorA(color, node.ParseInt32WithMinMax(transparencyAttr, 0, 100));
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

    public enum MarkerType
    {
        Circle,
        Square,
        Diamond,
        Triangle,
        X,
    }
    public enum LegendPosition
    {
        Top,
        Bottom,
        Left,
        Right,
        TopRight,
    }
    public enum DataLabelPosition
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
    public enum UnitType
    {
        Px,
        Cm,
        Pt,
        In,
    }
    public enum BlanksDisplayedAs
    {
        Span,
        Gap,
        Zero
    }
    public enum GroupingType
    {
        Standard,
        Clustered,
        Stacked,
        PercentStacked
    }
    public enum LayoutMode
    {
        Edge,
        Factor,
    }
    public enum LayoutTarget
    {
        Inner,
        Outer,
    }
    public enum ThreeDShape
    {
        Cone,
        ConeToMax,
        Box,
        Cylinder,
        Pyramid,
        PyramidToMaximum,
    }
    public enum PatternFillPreset
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
    public enum FillStyle
    {
        Auto,
        Solid,
        Pattern,
        Gradient,
        None
    }
    public enum ScatterStyle
    {
        Line,
        LineMarker,
        Marker,
        Smooth,
        SmoothMarker,
    }
    public enum RadarStyle
    {
        Standard,
        Marker,
        Filled,
    }
    public enum ShadowType
    {
        Inner,
        Outer,
        Preset,
    }
    public enum ShadowPreset
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
    public enum BevelPreset
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
    public enum MaterialPreset
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
    public enum LightingPreset
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
    public enum LightingDirection
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
    public enum DashPreset
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
    public enum CompoundLinePreset
    {
        Single,
        Double,
        ThickThin,
        ThinThick,
        Triple,
    }
    public enum LineCapPreset
    {
        Round,
        Square,
        Flat,
    }
    public enum LineJoinPreset
    {
        Round,
        Bevel,
        Miter,
    }
    public enum LineEndPreset
    {
        None,
        Triangle,
        Stealth,
        Diamond,
        Oval,
        Arrow,
    }
    public enum LineEndWidthPreset
    {
        Small,
        Medium,
        Large,
    }
    public enum LineEndLengthPreset
    {
        Small,
        Medium,
        Large,
    }
    public enum TextUnderlineStyle
    {
        None,
        Single,
        Double,
    }
    public enum TextStrikeStyle
    {
        None,
        Single,
        Double,
    }
    public enum TextCapsStyle
    {
        None,
        Small,
        All,
    }
    public enum HorizontalAlign
    {
        Left,
        Center,
        Right,
        Justify,
    }
    public enum VerticalAlign
    {
        Top,
        Center,
        Bottom,
    }
    public enum TextWrapPreset
    {
        None,
        Square,
    }

    internal interface ITitleOwner
    {
        TitleDefinition Title { get; set; }
    }
    internal interface IMarkerOwner
    {
        MarkerDefinition Marker { get; set; }
    }
    internal interface IValueLabelsContainer
    {
        ValueLabelDefinition ValueLabels { get; }
        ChartDefinition GetChartDefinition();
    }

    // Bridge tipi: dışarıdan sağlanan provider'ı taşır 
    public class ExcelDocumentAssets
    {
        public IChartLoadingProvider ChartLoadingProvider { get; set; }
    }
}

// Orijinal arabirimler (kullanıcının sağladığı) – değişmeden tutuldu
namespace OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents
{
    public interface IViewChartDefinition
    {
        // WordProcessing.Internal.ChartType yerinize doğrudan string/enum uyumlu kabul ediyoruz
        string ChartType { get; }

        string Title { get; }
        List<IViewChartSeriesDefinition> Series { get; set; }
    }
    public interface IViewChartSeriesDefinition
    {
        string Title { get; set; }
        string Color { set; get; }
        List<IViewChartPointDefinition> Points { get; set; }
    }
    public interface IViewChartPointDefinition
    {
        string Category { get; set; }
        double? Value { get; set; }
        double? X { get; set; }
        double? Y { get; set; }
    }

    public interface IChartLoadingProvider
    {
        IViewChartDefinition LoadViewChart(XmlNode node);
    }
}

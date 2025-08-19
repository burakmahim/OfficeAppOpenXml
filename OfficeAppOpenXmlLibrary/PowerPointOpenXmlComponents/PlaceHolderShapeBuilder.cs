using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAppOpenXmlLibrary.PowerPointOpenXmlComponents
{
    public class PlaceHolderShapeBuilder
    {
        public static Shape CreatePlaceholderShape(uint id, PlaceholderValues type, string text, string name, long x, long y, long cx, long cy, int fontSize = 24, bool bold = false, bool italic = false, string backgroundColor = "", string fontFamily = "Calibri", string textColor = "#000000", string zOrder = "front", uint? index = null)
        {
            fontSize = fontSize * 100;

            PlaceholderShape placeholderShape = new PlaceholderShape() { Type = type };

            if (index.HasValue)
                placeholderShape.Index = index.Value;

            Shape shape = new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = id, Name = name },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(placeholderShape)
                ),
                new ShapeProperties(
                    new A.Transform2D(
                        new A.Offset() { X = x, Y = y },
                        new A.Extents() { Cx = cx, Cy = cy }
                    ),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                ),
                new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.ParagraphProperties(),
                        new A.Run(
                            new A.RunProperties(
                                new A.SolidFill(new A.RgbColorModelHex() { Val = textColor.Replace("#", "") }),
                                new A.LatinFont() { Typeface = fontFamily }
                            )
                            {
                                FontSize = fontSize,
                                Bold = bold,
                                Italic = italic
                            },

                            new A.Text(text ?? " ")
                        )
                    )
                )
            );

            return shape;
        }
    }
}

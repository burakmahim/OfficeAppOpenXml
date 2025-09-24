using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAppOpenXmlLibrary.PowerPointOpenXmlComponents
{
    public class ShapeComponent
    {
        public static void  AddShape         (SlidePart slidePart, XElement shapeElement, ref uint shapeId)
        {
            // XML'den parametreleri al
            string shapeType = shapeElement.Attribute("type")?.Value ?? "Rectangle";
            string text = shapeElement.Attribute("text")?.Value ?? "";
            string fontFamily = shapeElement.Attribute("fontFamily")?.Value ?? "Arial";
            string textColor = shapeElement.Attribute("textColor")?.Value?.Replace("#", "") ?? "000000";
            string backgroundColor = shapeElement.Attribute("backgroundColor")?.Value?.Replace("#", "") ?? "";
            string borderColor = shapeElement.Attribute("borderColor")?.Value?.Replace("#", "") ?? "";
            string zOrder = shapeElement.Attribute("zOrder")?.Value ?? "front";

            // Koordinatlar
            (long x, long y, long cx, long cy) = CoordinatesParser.CoordinateParser(shapeElement, 5, 5, 10, 5);

            // Font özellikleri
            int fontSize = int.TryParse(shapeElement.Attribute("fontSize")?.Value, out int fs) ? fs : 18;
            bool bold = bool.TryParse(shapeElement.Attribute("bold")?.Value, out bool b) && b;
            bool italic = bool.TryParse(shapeElement.Attribute("italic")?.Value, out bool i) && i;
            
            // Border kalınlığı
            int borderWidth = int.TryParse(shapeElement.Attribute("borderWidth")?.Value, out int bw) ? bw : 1;

            // Shape oluştur
            Shape shape = CreateCustomShape(shapeId++, shapeType, text, x, y, cx, cy, 
                fontSize, bold, italic, fontFamily, textColor, backgroundColor, 
                borderColor, borderWidth);

            // Z-Order ayarla
            Z_OrderComponent.ZOrder(shape, slidePart.Slide.CommonSlideData.ShapeTree, zOrder);
        }

        public static Shape CreateCustomShape(uint id, string shapeType, string text, long x, long y, long cx, long cy, int fontSize = 18, bool bold = false, bool italic = false,  string fontFamily = "Arial", string textColor = "000000",  string backgroundColor = "", string borderColor = "", int borderWidth = 1)
        {
            // ShapeType'ı direkt parse etmeye çalış
            A.ShapeTypeValues geometryType = A.ShapeTypeValues.Rectangle; // default
            
            if (Enum.TryParse<A.ShapeTypeValues>(shapeType, true, out A.ShapeTypeValues parsedType))
            {
                geometryType = parsedType;
            }
            
            Shape shape = new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = id, Name = $"{shapeType} Shape" },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new ShapeProperties(
                    new A.Transform2D(
                        new A.Offset () { X  = x,  Y  = y },
                        new A.Extents() { Cx = cx, Cy = cy }
                    ),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = geometryType }
                )
            );

            // Arka plan rengi
            if (!string.IsNullOrWhiteSpace(backgroundColor))
            {
                shape.ShapeProperties.Append(
                    new A.SolidFill(
                        new A.RgbColorModelHex() { Val = backgroundColor }
                    )
                );
            }

            // Border (çizgi)
            if (!string.IsNullOrWhiteSpace(borderColor) || borderWidth > 0)
            {
                A.Outline outline = new A.Outline() { Width = borderWidth * 12700 }; // EMU cinsinden

                if (!string.IsNullOrWhiteSpace(borderColor))
                {
                    outline.Append(new A.SolidFill(
                        new A.RgbColorModelHex() { Val = borderColor }
                    ));
                }

                shape.ShapeProperties.Append(outline);
            }

            // Metin varsa ekle
            if (!string.IsNullOrWhiteSpace(text))
            {
                A.RunProperties runProps = new A.RunProperties(
                    new A.SolidFill(new A.RgbColorModelHex() { Val = textColor }),
                    new A.LatinFont() { Typeface = fontFamily }
                )
                {
                    FontSize = fontSize * 100,
                    Bold     = bold,
                    Italic   = italic
                };

                shape.TextBody = new TextBody(
                    new A.BodyProperties() 
                    { 
                        Wrap   = A.TextWrappingValues.Square,
                        Anchor = A.TextAnchoringTypeValues.Center
                    },
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center },
                        new A.Run(runProps, new A.Text(text))
                    )
                );
            }

            return shape;
        }
    }
}

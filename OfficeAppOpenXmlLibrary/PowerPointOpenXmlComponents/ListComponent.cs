using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAppOpenXmlLibrary.PowerPointOpenXmlComponents
{
    public class ListComponent
    {
        public static void  AddList(SlidePart slidePart, XElement listElement, ref uint shapeId)
        {
            List<XElement> items = listElement.Elements("item").ToList();

            (long x, long y, long cx, long cy) = CoordinatesParser.CoordinateParser(listElement, 2.54, 4.45, 20.32, 12.7);

            string fontFamily       = listElement.Attribute("fontFamily")?.Value ?? "Arial";
            string textColor        = listElement.Attribute("textColor")?.Value?.Replace("#", "") ?? "000000";
            string backgroundColor  = listElement.Attribute("backgroundColor")?.Value?.Replace("#", "") ?? "";
            string zOrder           = listElement.Attribute("zOrder")?.Value ?? "front";

            bool bold   = false;
            bool italic = false;

            bool.TryParse(listElement.Attribute("bold"  )?.Value, out bold);
            bool.TryParse(listElement.Attribute("italic")?.Value, out italic);

            int fontSize = int.TryParse(listElement.Attribute("fontSize")?.Value, out int parsedFontSize) ? parsedFontSize : 24;

            Shape shape = CreateListShape(shapeId++, "List Placeholder", items, x, y, cx, cy, fontSize, bold, italic, backgroundColor, fontFamily, textColor );

            Z_OrderComponent.ZOrder(shape, slidePart.Slide.CommonSlideData.ShapeTree, zOrder);
        }

        public static Shape CreateListShape(uint id, string name, List<XElement> items, long x, long y, long cx, long cy, int fontSize = 24, bool bold = false, bool italic = false, string backgroundColor = "", string fontFamily = "Arial", string textColor = "000000")
        {
            fontSize = fontSize * 100;

            Shape listShape = new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = id, Name = name },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()
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
                    new A.ListStyle()
                )
            );

            if (!string.IsNullOrWhiteSpace(backgroundColor))
            {
                listShape.ShapeProperties.Append(
                    new A.SolidFill(
                        new A.RgbColorModelHex() { Val = backgroundColor.Replace("#", "") }
                    )
                );
            }

            List<A.Paragraph> paragraphs = new List<A.Paragraph>();

            foreach (XElement itemNode in items)
            {
                string text = itemNode.Value;

                int itemFontSize = int.TryParse(itemNode.Attribute("fontSize")?.Value, out int fs) ? fs * 100 : fontSize;
                bool itemBold = bool.TryParse(itemNode.Attribute("bold")?.Value, out bool b) && b;
                bool itemItalic = bool.TryParse(itemNode.Attribute("italic")?.Value, out bool i) && i;
                string itemFontFamily = itemNode.Attribute("fontFamily")?.Value ?? fontFamily;
                string itemTextColor = itemNode.Attribute("textColor")?.Value?.Replace("#", "") ?? textColor;

                A.Paragraph paragraph = new A.Paragraph(
                    new A.ParagraphProperties(
                        new A.BulletFont() { Typeface = "Arial" },
                        new A.CharacterBullet() { Char = "•" }
                    )
                    {
                        Level = 0
                    },
                    new A.Run(
                        new A.RunProperties(
                            new A.SolidFill(new A.RgbColorModelHex() { Val = itemTextColor.Replace("#", "") }),
                            new A.LatinFont() { Typeface = itemFontFamily }
                        )
                        {
                            FontSize = itemFontSize,
                            Bold = itemBold,
                            Italic = itemItalic
                        },
                        new A.Text(text)
                    )
                );
                paragraphs.Add(paragraph);
            }

            listShape.TextBody.Append(paragraphs.ToArray());

            return listShape;
        }
    }
}

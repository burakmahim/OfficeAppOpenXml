using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Xml.Linq;

namespace OfficeAppOpenXmlLibrary.PowerPointOpenXmlComponents
{
    public class TextBoxComponent
    {
        public static void AddTextBox(SlidePart slidePart, XElement textboxNode, ref uint shapeId)
        {
            string textboxText = textboxNode.Value ?? "";

            (long x, long y, long cx, long cy) = CoordinatesParser.CoordinateParser(textboxNode, 1, 1, 15, 15);

            string fontFamily       = textboxNode.Attribute("fontFamily")?.Value ?? "Arial";
            string textColor        = textboxNode.Attribute("textColor")?.Value?.Replace("#", "") ?? "000000";
            string backgroundColor  = textboxNode.Attribute("backgroundColor")?.Value?.Replace("#", "") ?? "";

            bool bold               = bool.Parse(textboxNode.Attribute("bold")?.Value ?? "false");
            bool italic             = bool.Parse(textboxNode.Attribute("italic")?.Value ?? "false");

            int fontSize            = int.Parse(textboxNode.Attribute("fontSize")?.Value ?? "20");

            string zOrder           = textboxNode.Attribute("zOrder")?.Value ?? "front";

            Shape shape             = PlaceHolderShapeBuilder.CreatePlaceholderShape(shapeId++, PlaceholderValues.Body, textboxText, "Textbox Placeholder", x, y, cx, cy, fontSize, bold, italic, backgroundColor, fontFamily, textColor);
            Z_OrderComponent.ZOrder(shape, slidePart.Slide.CommonSlideData.ShapeTree, zOrder);
        }
    }
}

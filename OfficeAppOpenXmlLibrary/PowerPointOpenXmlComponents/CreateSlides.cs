using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAppOpenXmlLibrary.PowerPointOpenXmlComponents
{
    public class CreateSlides
    {
        public static void CreateDefaultSlide(SlidePart slidePart, XElement slideNode, ref uint shapeId)
        {
            ShapeTree shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;

            string title = slideNode.Element("title")?.Value ?? "";
            string content = slideNode.Element("content")?.Value ?? "";

            foreach (Shape shape in shapeTree.Descendants<Shape>())
            {
                PlaceholderShape? placeholder = shape.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?
                    .GetFirstChild<PlaceholderShape>();

                if (placeholder != null && placeholder.Type == PlaceholderValues.Title)
                {
                    shape.TextBody = new TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.Run(new A.Text(title))),
                            new A.ParagraphProperties()
                    );
                }

                if (placeholder != null && placeholder.Type == PlaceholderValues.Body)
                {
                    shape.TextBody = new TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(new A.Run(new A.Text(content)))
                    );
                }

            }
        }
        public static void CreateHeaderSlide(SlidePart slidePart, XElement slideNode, ref uint shapeId)
        {
            ShapeTree shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;

            string title = slideNode.Element("title")?.Value ?? "";
            string subtitle = slideNode.Element("subtitle")?.Value ?? "";

            foreach (Shape shape in shapeTree.Descendants<Shape>())
            {
                PlaceholderShape? placeholder = shape.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?
                    .GetFirstChild<PlaceholderShape>();

                if (placeholder != null && placeholder.Type == PlaceholderValues.Title)
                {
                    shape.TextBody = new TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(new A.Run(new A.Text(title)))
                    );
                }

                if (placeholder != null && placeholder.Type == PlaceholderValues.SubTitle)
                {
                    shape.TextBody = new TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(new A.Run(new A.Text(subtitle)))
                    );
                }
            }
        }
        public static void CreateTwoContentSlide(SlidePart slidePart, XElement slideNode, ref uint shapeId)
        {
            ShapeTree? shapeTree = slidePart?.Slide?.CommonSlideData?.ShapeTree;

            string title = slideNode.Element("title")?.Value ?? "";
            string leftContent = slideNode.Element("leftContent")?.Value ?? "";
            string rightContent = slideNode.Element("rightContent")?.Value ?? "";

            List<Shape> bodyShapes = shapeTree?.Descendants<Shape>()
                .Where(s => s.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?
                    .GetFirstChild<PlaceholderShape>()?.Type == PlaceholderValues.Body)
                .ToList();


            foreach (Shape shape in shapeTree.Descendants<Shape>())
            {
                PlaceholderShape? placeholder = shape.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?
                    .GetFirstChild<PlaceholderShape>();

                if (placeholder == null || placeholder.Type == null)
                    continue;


                if (placeholder.Type == PlaceholderValues.Title)
                {
                    shape.TextBody = new TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(new A.Run(new A.Text(title)))
                    );
                }
            }

            if (bodyShapes.Count > 0)
            {
                bodyShapes[0].TextBody = new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(leftContent)))
                );
            }

            if (bodyShapes.Count > 1)
            {
                bodyShapes[1].TextBody = new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(rightContent)))
                );
            }

        }
        public static void CreateEmptySlide(SlidePart slidePart, ref uint shapeId)
        {
            ShapeTree? shapeTree = slidePart?.Slide?.CommonSlideData?.ShapeTree;
        }
    }
}

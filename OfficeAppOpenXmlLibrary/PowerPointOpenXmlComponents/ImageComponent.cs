using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAppOpenXmlLibrary.PowerPointOpenXmlComponents
{
    public class ImageComponent
    {
        public static void AddImage   (SlidePart slidePart, XElement imageNode, ref uint shapeId)
        {
            string? path = imageNode?.Attribute("path")?.Value;
            string? zOrder = imageNode?.Attribute("zOrder")?.Value ?? "front";

            if (path != null)
            {

                (long x, long y, long cx, long cy) = CoordinatesParser.CoordinateParser(imageNode, 19.44, 0.83, 4.17, 4.17);

                if (path.StartsWith("data:image"))
                {
                    string base64Data = path.Substring(path.IndexOf(",") + 1);
                    byte[] imageBytes = Convert.FromBase64String(base64Data);

                    using (MemoryStream memStream = new MemoryStream(imageBytes))
                    {
                        ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Jpeg);
                        imagePart.FeedData(memStream);

                        string imagePartId = slidePart.GetIdOfPart(imagePart);

                        Picture picture = new Picture(
                            new NonVisualPictureProperties(
                                new NonVisualDrawingProperties() { Id = shapeId++, Name = "Picture " + shapeId },
                                new NonVisualPictureDrawingProperties(new A.PictureLocks() { NoChangeAspect = true }),
                                new ApplicationNonVisualDrawingProperties()
                            ),
                            new BlipFill(
                                new A.Blip() { Embed = imagePartId },
                                new A.Stretch(new A.FillRectangle())
                            ),
                            new ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset() { X = x, Y = y },
                                    new A.Extents() { Cx = cx, Cy = cy }
                                ),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                            )
                        );

                        ZOrderImage(picture, slidePart.Slide.CommonSlideData.ShapeTree, zOrder);
                    }
                }
                else if (path.StartsWith("http://") || path.StartsWith("https://"))
                {
                    using (var webClient = new System.Net.WebClient())
                    {
                        byte[] imageBytes = webClient.DownloadData(path);

                        using (MemoryStream memStream = new MemoryStream(imageBytes))
                        {
                            ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Jpeg);
                            imagePart.FeedData(memStream);

                            string imagePartId = slidePart.GetIdOfPart(imagePart);

                            Picture picture = new Picture(
                                new NonVisualPictureProperties(
                                    new NonVisualDrawingProperties() { Id = shapeId++, Name = "Picture " + shapeId },
                                    new NonVisualPictureDrawingProperties(new A.PictureLocks() { NoChangeAspect = true }),
                                    new ApplicationNonVisualDrawingProperties()
                                ),
                                new BlipFill(
                                    new A.Blip() { Embed = imagePartId },
                                    new A.Stretch(new A.FillRectangle())
                                ),
                                new ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset() { X = x, Y = y },
                                        new A.Extents() { Cx = cx, Cy = cy }
                                    ),
                                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                                )
                            );

                            ZOrderImage(picture, slidePart.Slide.CommonSlideData.ShapeTree, zOrder);
                        }
                    }
                }
                else
                {

                    if (System.IO.File.Exists(path))
                    {
                        using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read))
                        {
                            ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Jpeg);
                            imagePart.FeedData(fileStream);

                            string imagePartId = slidePart.GetIdOfPart(imagePart);

                            Picture picture = new Picture(
                                new NonVisualPictureProperties(
                                    new NonVisualDrawingProperties() { Id = shapeId++, Name = "Picture " + shapeId },
                                    new NonVisualPictureDrawingProperties(new A.PictureLocks() { NoChangeAspect = true }),
                                    new ApplicationNonVisualDrawingProperties()
                                ),
                                new BlipFill(
                                    new A.Blip() { Embed = imagePartId },
                                    new A.Stretch(new A.FillRectangle())
                                ),
                                new ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset() { X = x, Y = y },
                                        new A.Extents() { Cx = cx, Cy = cy }
                                    ),
                                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                                )
                            );

                            ZOrderImage(picture, slidePart.Slide.CommonSlideData.ShapeTree, zOrder);
                        }
                    }
                }

            }

        }
        public static void ZOrderImage(Picture picture, ShapeTree shapeTree, string zOrder)      
        {
            if (zOrder == "back")
            {
                shapeTree.InsertAt(picture, 0);
            }
            else if (zOrder == "front")
            {
                shapeTree.Append(picture);
            }
        }
    }
}

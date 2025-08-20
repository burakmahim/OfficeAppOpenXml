using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeAppOpenXmlLibrary.PowerPointOpenXmlComponents;
using System.Xml.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAppOpenXmlLibrary
{
    public class PowerPointLibrary
    {
        public static byte[] CreatePowerPointPresentation(string xmlContent)
        {

            using (MemoryStream memoryStream = new MemoryStream())
            {
                using (PresentationDocument presentationDoc = PresentationDocument.Create(memoryStream, PresentationDocumentType.Presentation, true))
                {
                    PresentationPart presentationPart = presentationDoc.AddPresentationPart();
                    presentationPart.Presentation = new Presentation();

                    SlideMasterPart slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();

                    SlideLayoutIdList slideLayoutIdList = new SlideLayoutIdList();
                    uint slideLayoutId = 1;

                    var slideLayoutParts = new Dictionary<string, SlideLayoutPart>();

                    uint shapeId = 1;

                    XElement presentationXml = XElement.Parse(xmlContent);

                    SlideLayoutPart defaultLayoutPart = LayoutComponent.CreateDefaultLayout(slideMasterPart, ref shapeId);
                    slideLayoutParts["default"] = defaultLayoutPart;
                    slideLayoutIdList.Append(new SlideLayoutId()
                    {
                        Id = slideLayoutId++,
                        RelationshipId = slideMasterPart.GetIdOfPart(defaultLayoutPart)
                    });

                    SlideLayoutPart headerLayoutPart = LayoutComponent.CreateHeaderLayout(slideMasterPart, ref shapeId);
                    slideLayoutParts["header"] = headerLayoutPart;
                    slideLayoutIdList.Append(new SlideLayoutId()
                    {
                        Id = slideLayoutId++,
                        RelationshipId = slideMasterPart.GetIdOfPart(headerLayoutPart)
                    });

                    SlideLayoutPart twoContentLayoutPart = LayoutComponent.CreateTwoContentLayout(slideMasterPart, ref shapeId);
                    slideLayoutParts["twoContent"] = twoContentLayoutPart;
                    slideLayoutIdList.Append(new SlideLayoutId()
                    {
                        Id = slideLayoutId++,
                        RelationshipId = slideMasterPart.GetIdOfPart(twoContentLayoutPart)
                    });

                    SlideLayoutPart emptyLayoutPart = LayoutComponent.CreateBlankLayout(slideMasterPart, ref shapeId);
                    slideLayoutParts["empty"] = emptyLayoutPart;
                    slideLayoutIdList.Append(new SlideLayoutId()
                    {
                        Id = slideLayoutId++,
                        RelationshipId = slideMasterPart.GetIdOfPart(emptyLayoutPart)
                    });

                    ShapeTree slideMasterShapeTree = new ShapeTree(
                        new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties() { Id = shapeId++, Name = "MasterShapeTree" },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()
                        ),
                        new GroupShapeProperties()
                    );

                    Shape slideNumberShape = PlaceHolderShapeBuilder.CreatePlaceholderShape(
                    shapeId++, PlaceholderValues.SlideNumber, "Slide Number", "Slide Number", 914400, 6368400, 4114800, 363600, 12, false, false, "", "Calibri", "#7E7E7E"
                    );

                    slideMasterShapeTree.AppendChild(slideNumberShape);

                    slideMasterPart.SlideMaster = new SlideMaster(
                        new CommonSlideData(slideMasterShapeTree),
                        slideLayoutIdList,
                        new TextStyles(
                            new TitleStyle(
                                new A.Level1ParagraphProperties(
                                new A.DefaultRunProperties(new A.LatinFont() { Typeface = "Bradley Hand ITC" })
                                {
                                    FontSize = 3600,
                                    Bold = true,
                                    Italic = true,
                                }
                                )
                            ),
                            new BodyStyle(
                                new A.Level1ParagraphProperties(
                                    new A.DefaultRunProperties()
                                    {
                                        FontSize = 2400,
                                        Bold = false,
                                        Italic = false,
                                    }
                                )
                            ),
                            new OtherStyle()
                        ),
                        new HeaderFooter() { Footer = true, DateTime = true, SlideNumber = true }
                    );

                    slideMasterPart.SlideMaster.Save();

                    string masterRelId = presentationPart.GetIdOfPart(slideMasterPart);

                    presentationPart.Presentation.Append(
                        new SlideMasterIdList(
                            new SlideMasterId()
                            {
                                Id = 1U,
                                RelationshipId = masterRelId
                            })
                    );

                    SlideIdList slideIdList = new SlideIdList();
                    uint slideId = 256;

                    XElement? footerNode = presentationXml.Element("footer");
                    string footer = footerNode?.Value ?? "";


                    foreach (XElement slideNode in presentationXml.Elements("slide"))
                    {
                        SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

                        string layoutType = slideNode.Attribute("layout")?.Value ?? "default";
                        string addFooterAttribute = slideNode.Attribute("addFooter")?.Value ?? "false";
                        bool addFooter;
                        bool.TryParse(addFooterAttribute, out addFooter);

                        if (layoutType == "default")
                        {
                            slidePart.Slide = new Slide(new CommonSlideData(
                                (ShapeTree)defaultLayoutPart.SlideLayout.CommonSlideData.ShapeTree.CloneNode(true)
                            ));
                            slidePart.AddPart(defaultLayoutPart);

                            CreateSlides.CreateDefaultSlide(slidePart, slideNode, ref shapeId);
                        }
                        else if (layoutType == "header")
                        {

                            slidePart.Slide = new Slide(new CommonSlideData(
                                 (ShapeTree)headerLayoutPart.SlideLayout.CommonSlideData.ShapeTree.CloneNode(true)

                            ));
                            slidePart.AddPart(headerLayoutPart);

                            CreateSlides.CreateHeaderSlide(slidePart, slideNode, ref shapeId);

                        }
                        else if (layoutType == "twoContent")
                        {
                            slidePart.Slide = new Slide(new CommonSlideData(
                                 (ShapeTree)twoContentLayoutPart.SlideLayout.CommonSlideData.ShapeTree.CloneNode(true)
                            ));
                            slidePart.AddPart(twoContentLayoutPart);

                            CreateSlides.CreateTwoContentSlide(slidePart, slideNode, ref shapeId);
                        }
                        else if (layoutType == "empty")
                        {
                            slidePart.Slide = new Slide(new CommonSlideData(
                                 (ShapeTree)emptyLayoutPart.SlideLayout.CommonSlideData.ShapeTree.CloneNode(true)
                            ));
                            slidePart.AddPart(emptyLayoutPart);

                            CreateSlides.CreateEmptySlide(slidePart, ref shapeId);
                        }
                        else
                        {
                            slidePart.Slide = new Slide(new CommonSlideData(
                                (ShapeTree)defaultLayoutPart.SlideLayout.CommonSlideData.ShapeTree.CloneNode(true)
                            ));

                            slidePart.AddPart(defaultLayoutPart);

                        }

                        slideIdList.Append(new SlideId()
                        {
                            Id = slideId++,
                            RelationshipId = presentationPart.GetIdOfPart(slidePart)
                        });

                        string? backgroundImagePath = slideNode.Attribute("backgroundImage")?.Value;

                        string? slideBackgroundColor = slideNode.Attribute("slideBackgroundColor")?.Value;


                        IEnumerable<XElement> imageElements = slideNode.Elements("image");
                        if (imageElements!= null)
                        {
                            foreach (XElement image in imageElements)
                            {
                                ImageComponent.AddImage(slidePart, image, ref shapeId);
                            }
                        }


                        IEnumerable<XElement> textboxElements = slideNode.Elements("textbox");
                        if (textboxElements != null)
                        {
                            foreach (XElement textbox in textboxElements)
                            {
                                TextBoxComponent.AddTextBox(slidePart, textbox, ref shapeId);
                            }
                        }

                        IEnumerable<XElement> listElements = slideNode.Elements("list");
                        if (listElements != null)
                        {
                            foreach (XElement list in listElements)
                            {
                                ListComponent.AddList(slidePart, list, ref shapeId);
                            }
                        }

                    }

                    presentationPart.Presentation.Append(slideIdList);
                    presentationPart.Presentation.Save();
                }

                return memoryStream.ToArray();
            }

        }

    }
}

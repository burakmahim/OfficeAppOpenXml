using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using OfficeAppOpenXmlLibrary.OpenXmlChartComponent;
using OfficeAppOpenXmlLibrary.PowerPointOpenXmlComponents;

using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeAppOpenXmlLibrary
{
    public class PowerPointLibrary
    {
        /// <summary>
        /// DSL (pptx.xml) içeriğini kullanarak bir PPTX oluşturur ve
        /// aynı DSL'i dosyaya CustomXmlPart olarak gömer (round-trip için).
        /// </summary>
        public static byte[] CreatePowerPointPresentation(string xmlContent)
        {
            using (var memoryStream = new MemoryStream())
            {
                using (var presentationDoc = PresentationDocument.Create(
                           memoryStream,
                           PresentationDocumentType.Presentation,
                           autoSave: true))
                {
                    // Root presentation part
                    var presentationPart = presentationDoc.AddPresentationPart();
                    presentationPart.Presentation = new Presentation();

                    // Slide master + layout'lar
                    var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
                    var slideLayoutIdList = new SlideLayoutIdList();
                    uint slideLayoutId = 1;

                    var slideLayoutParts = new Dictionary<string, SlideLayoutPart>(StringComparer.OrdinalIgnoreCase);

                    uint shapeId = 1; // tüm şekiller için global artan kimlik

                    // DSL'i parse et
                    var presentationXml = XElement.Parse(xmlContent);

                    // Layout'lar
                    var defaultLayoutPart = LayoutComponent.CreateDefaultLayout(slideMasterPart, ref shapeId);
                    slideLayoutParts["default"] = defaultLayoutPart;
                    slideLayoutIdList.Append(new SlideLayoutId
                    {
                        Id = slideLayoutId++,
                        RelationshipId = slideMasterPart.GetIdOfPart(defaultLayoutPart)
                    });

                    var headerLayoutPart = LayoutComponent.CreateHeaderLayout(slideMasterPart, ref shapeId);
                    slideLayoutParts["header"] = headerLayoutPart;
                    slideLayoutIdList.Append(new SlideLayoutId
                    {
                        Id = slideLayoutId++,
                        RelationshipId = slideMasterPart.GetIdOfPart(headerLayoutPart)
                    });

                    var twoContentLayoutPart = LayoutComponent.CreateTwoContentLayout(slideMasterPart, ref shapeId);
                    slideLayoutParts["twoContent"] = twoContentLayoutPart;
                    slideLayoutIdList.Append(new SlideLayoutId
                    {
                        Id = slideLayoutId++,
                        RelationshipId = slideMasterPart.GetIdOfPart(twoContentLayoutPart)
                    });

                    var emptyLayoutPart = LayoutComponent.CreateBlankLayout(slideMasterPart, ref shapeId);
                    slideLayoutParts["empty"] = emptyLayoutPart;
                    slideLayoutIdList.Append(new SlideLayoutId
                    {
                        Id = slideLayoutId++,
                        RelationshipId = slideMasterPart.GetIdOfPart(emptyLayoutPart)
                    });

                    // Master'ın kendi shape tree'si ve slayt numarası placeholder'ı
                    var slideMasterShapeTree = new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = shapeId++, Name = "MasterShapeTree" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()
                        ),
                        new GroupShapeProperties()
                    );

                    // Slide number placeholder (P.Shape olarak açık ad)
                    P.Shape slideNumberShape = PlaceHolderShapeBuilder.CreatePlaceholderShape(
                        shapeId++,
                        PlaceholderValues.SlideNumber,
                        "Slide Number",
                        "Slide Number",
                        // x, y, cx, cy (EMU)
                        914400, 6368400, 4114800, 363600,
                        // font size pt, bold/italic, font family, color
                        12,
                        false,
                        false,
                        "",
                        "Calibri",
                        "#7E7E7E"
                    );

                    slideMasterShapeTree.AppendChild(slideNumberShape);

                    // Master + text styles + footer
                    slideMasterPart.SlideMaster = new SlideMaster(
                        new CommonSlideData(slideMasterShapeTree),
                        slideLayoutIdList,
                        new TextStyles(
                            new TitleStyle(
                                new A.Level1ParagraphProperties(
                                    new A.DefaultRunProperties(new A.LatinFont { Typeface = "Bradley Hand ITC" })
                                    {
                                        FontSize = 3600,
                                        Bold = true,
                                        Italic = true
                                    }
                                )
                            ),
                            new BodyStyle(
                                new A.Level1ParagraphProperties(
                                    new A.DefaultRunProperties
                                    {
                                        FontSize = 2400,
                                        Bold = false,
                                        Italic = false
                                    }
                                )
                            ),
                            new OtherStyle()
                        ),
                        new HeaderFooter { Footer = true, DateTime = true, SlideNumber = true }
                    );

                    slideMasterPart.SlideMaster.Save();

                    // Master'ı presentation'a bağla
                    string masterRelId = presentationPart.GetIdOfPart(slideMasterPart);
                    presentationPart.Presentation.Append(new SlideMasterIdList(
                        new SlideMasterId
                        {
                            Id = 1U,
                            RelationshipId = masterRelId
                        }
                    ));

                    // Slides
                    var slideIdList = new SlideIdList();
                    uint slideId = 256;

                    // (Ops.) Genel footer XML düğümü (şimdilik kullanılmıyor)
                    // var footerNode = presentationXml.Element("footer");
                    // var footer     = footerNode?.Value ?? string.Empty;

                    foreach (var slideNode in presentationXml.Elements("slide"))
                    {
                        var slidePart = presentationPart.AddNewPart<SlidePart>();

                        string layoutType = slideNode.Attribute("layout")?.Value ?? "default";
                        string addFooterStr = slideNode.Attribute("addFooter")?.Value ?? "false";
                        bool addFooter = bool.TryParse(addFooterStr, out var af) && af;

                        // Belirlenen layout'a göre slide oluştur
                        if (layoutType.Equals("default", StringComparison.OrdinalIgnoreCase))
                        {
                            slidePart.Slide = new Slide(new CommonSlideData(
                                (ShapeTree)slideLayoutParts["default"].SlideLayout.CommonSlideData.ShapeTree.CloneNode(true)
                            ));
                            slidePart.AddPart(slideLayoutParts["default"]);
                            CreateSlides.CreateDefaultSlide(slidePart, slideNode, ref shapeId);
                        }
                        else if (layoutType.Equals("header", StringComparison.OrdinalIgnoreCase))
                        {
                            slidePart.Slide = new Slide(new CommonSlideData(
                                (ShapeTree)slideLayoutParts["header"].SlideLayout.CommonSlideData.ShapeTree.CloneNode(true)
                            ));
                            slidePart.AddPart(slideLayoutParts["header"]);
                            CreateSlides.CreateHeaderSlide(slidePart, slideNode, ref shapeId);
                        }
                        else if (layoutType.Equals("twoContent", StringComparison.OrdinalIgnoreCase))
                        {
                            slidePart.Slide = new Slide(new CommonSlideData(
                                (ShapeTree)slideLayoutParts["twoContent"].SlideLayout.CommonSlideData.ShapeTree.CloneNode(true)
                            ));
                            slidePart.AddPart(slideLayoutParts["twoContent"]);
                            CreateSlides.CreateTwoContentSlide(slidePart, slideNode, ref shapeId);
                        }
                        else if (layoutType.Equals("empty", StringComparison.OrdinalIgnoreCase))
                        {
                            slidePart.Slide = new Slide(new CommonSlideData(
                                (ShapeTree)slideLayoutParts["empty"].SlideLayout.CommonSlideData.ShapeTree.CloneNode(true)
                            ));
                            slidePart.AddPart(slideLayoutParts["empty"]);
                            CreateSlides.CreateEmptySlide(slidePart, ref shapeId);
                        }
                        else
                        {
                            // bilinmeyen layout -> default
                            slidePart.Slide = new Slide(new CommonSlideData(
                                (ShapeTree)slideLayoutParts["default"].SlideLayout.CommonSlideData.ShapeTree.CloneNode(true)
                            ));
                            slidePart.AddPart(slideLayoutParts["default"]);
                        }

                        // Slide ID listesine ekle
                        slideIdList.Append(new SlideId
                        {
                            Id = slideId++,
                            RelationshipId = presentationPart.GetIdOfPart(slidePart)
                        });

                        // İçerikler
                        foreach (var img in slideNode.Elements("image"))
                            ImageComponent.AddImage(slidePart, img, ref shapeId);

                        foreach (var tb in slideNode.Elements("textbox"))
                            TextBoxComponent.AddTextBox(slidePart, tb, ref shapeId);

                        foreach (var list in slideNode.Elements("list"))
                            ListComponent.AddList(slidePart, list, ref shapeId);

                        foreach (var chart in slideNode.Elements("chart"))
                            ChartComponent.PowerPointAddChart(slidePart, chart, ref shapeId);

                        _ = addFooter; // footer istenirse burada uygulanabilir
                    }

                    // Slides'i presentation'a bağla ve kaydet
                    presentationPart.Presentation.Append(slideIdList);
                    presentationPart.Presentation.Save();

                    // >>> DSL'i PPTX içine CustomXmlPart olarak göm (ROUND-TRIP) <<<
                    var cx = presentationPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                    using (var sw = new StreamWriter(cx.GetStream(FileMode.Create, FileAccess.Write)))
                    {
                        sw.Write(xmlContent);
                    }
                    // <<< gömme bitti
                }

                // MemoryStream to byte[]
                return memoryStream.ToArray();
            }
        }
    }
}

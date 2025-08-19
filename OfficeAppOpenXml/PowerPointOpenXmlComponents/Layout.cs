using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Text;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAppOpenXmlLibrary.PowerPointOpenXmlComponents
{
    class Layout
    {
        public SlideLayoutPart CreateBlankLayout(SlideMasterPart slideMasterPart, ref uint shapeId)
        {
            SlideLayoutPart slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();

            ShapeTree blankShapeTree = new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties() { Id = shapeId++, Name = "Blank ShapeTree" },
                    new NonVisualGroupShapeProperties(),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new GroupShapeProperties()
            );

            SlideLayout slideLayout = new SlideLayout(
                new CommonSlideData(blankShapeTree),
                new A.ColorMap()
            )
            { Type = SlideLayoutValues.Blank };

            slideLayout.CommonSlideData.Name = "Boş İçerik";
            slideLayout.Type = SlideLayoutValues.Blank;
            slideLayout.Preserve = true;
            slideLayout.ShowMasterShapes = true;

            slideLayoutPart.SlideLayout = slideLayout;
            slideLayoutPart.SlideLayout.Save();

            return slideLayoutPart;
        }

        public SlideLayoutPart CreateDefaultLayout(SlideMasterPart slideMasterPart, ref uint shapeId)
        {
            SlideLayoutPart slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();

            ShapeTree defaultShapeTree = new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties() { Id = shapeId++, Name = "Default ShapeTree" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new GroupShapeProperties()
            )
                ;

            defaultShapeTree.Append(CreatePlaceholderShape(
                shapeId++, PlaceholderValues.Title, text: "", "Title", 914400, 457200, 7315200, 1000000,
                fontSize: 32, bold: true
            ));

            defaultShapeTree.Append(CreatePlaceholderShape(
                shapeId++, PlaceholderValues.Body, text: "", "Content", 914400, 1600200, 7315200, 4572000
            ));

            defaultShapeTree.Append(CreatePlaceholderShape(
                shapeId++, PlaceholderValues.Footer, text: "", "Footer", 4039200, 6368400, 4114800, 363600, 12, false, false, "", "Calibri", "#7E7E7E"
            ));

            SlideLayout slideLayout = new SlideLayout(
                new CommonSlideData(defaultShapeTree),
                new A.ColorMap()
            )
            { Type = SlideLayoutValues.VerticalTitleAndText };

            slideLayout.CommonSlideData.Name = "Başlık ve İçerik";
            slideLayout.Type = SlideLayoutValues.VerticalTitleAndText;
            slideLayout.Preserve = true;
            slideLayout.ShowMasterShapes = true;

            slideLayoutPart.SlideLayout = slideLayout;
            slideLayoutPart.SlideLayout.Save();

            return slideLayoutPart;
        }
    }
}

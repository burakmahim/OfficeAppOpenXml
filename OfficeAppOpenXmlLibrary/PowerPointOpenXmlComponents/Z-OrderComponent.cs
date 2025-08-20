using DocumentFormat.OpenXml.Presentation;

namespace OfficeAppOpenXmlLibrary.PowerPointOpenXmlComponents
{
    public class Z_OrderComponent
    {
        public static void ZOrder(Shape shape, ShapeTree shapeTree, string zOrder = "")
        {
            if (zOrder == "back")
            {
                shapeTree.InsertAt(shape, 0);
            }

            else if (zOrder == "front")
            {
                shapeTree.Append(shape);
            }
        }
    }
}

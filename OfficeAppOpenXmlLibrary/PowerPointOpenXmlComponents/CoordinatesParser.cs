using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OfficeAppOpenXmlLibrary.PowerPointOpenXmlComponents
{
    public class CoordinatesParser
    {
        public static (long x, long y, long width, long height) CoordinateParser(XElement element, double defX, double defY, double defCx, double defCy)
        {
            const long emuConstant = 360000;

            long x      = (long)((double.TryParse(element.Attribute("x" )?.Value.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsedX ) ? parsedX : defX ) * emuConstant);
            long y      = (long)((double.TryParse(element.Attribute("y" )?.Value.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsedY ) ? parsedY : defY ) * emuConstant);
            long width  = (long)((double.TryParse(element.Attribute("cx")?.Value.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsedCx) ? parsedCx: defCx) * emuConstant);
            long height = (long)((double.TryParse(element.Attribute("cy")?.Value.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsedCy) ? parsedCy: defCy) * emuConstant);

            return (x, y, width, height);
        }
    }
}

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Presentation;
using System.Xml;
using OfficeAppOpenXmlLibrary.PowerPointOpenXmlComponents;
using System.Xml.Linq;

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
                    XElement presentationXml = XElement.Parse(xmlContent);

                    PresentationPart presentationPart = presentationDoc.AddPresentationPart();
                    presentationPart.Presentation = new Presentation();

                    SlideMasterPart slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();

                    SlideLayoutIdList slideLayoutIdList = new SlideLayoutIdList();
                    uint slideLayoutId = 1;
                    uint shapeId = 1;

                    foreach (XElement slideNode in presentationXml.Elements("slide"))
                    {

                    }


                    presentationDoc.Save();
                    return memoryStream.ToArray();
                }
            }
        }
    }
}
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Presentation;
using OfficeAppOpenXmlLibrary.PowerPointOpenXmlComponents;

namespace OfficeAppOpenXmlLibrary
{
    public class PowerPointLibrary
    {
        public void CreatePowerPointPresentation()
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {
                using (PresentationDocument presentationDoc = PresentationDocument.Create(memoryStream, PresentationDocumentType.Presentation, true))
                {
                    PresentationPart presentationPart = presentationDoc.AddPresentationPart();
                    presentationPart.Presentation = new Presentation();

                    SlideMasterPart slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();        

                }
            }
        }
    }
}

using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeAppOpenXmlLibrary.ExcelOpenXmlComponents
{
    public static class ExcelChartHelper
    {
        public static void AddChart(WorksheetPart wsPart, ChartSpec spec, string topLeftA1, string bottomRightA1)
        {
            if (wsPart == null) throw new ArgumentNullException(nameof(wsPart));
            if (spec == null) throw new ArgumentNullException(nameof(spec));
            if (wsPart.Worksheet == null)
                throw new InvalidOperationException("WorksheetPart.Worksheet null. AddChart'tan önce worksheetPart.Worksheet = worksheet; yapın.");

            // 1) DrawingPart + WorksheetDrawing
            var drawingsPart = wsPart.DrawingsPart ?? wsPart.AddNewPart<DrawingsPart>();
            if (drawingsPart.WorksheetDrawing == null)
                drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();

            // 2) ChartPart + ChartSpace
            var chartPart = drawingsPart.AddNewPart<ChartPart>();
            var chartSpace = ChartXmlBuilder.BuildChartSpace(spec)
                              ?? throw new InvalidOperationException("ChartSpace null.");
            chartPart.ChartSpace = chartSpace;
            chartPart.ChartSpace.Save();

            // 3) Anchor (iki hücre arası)
            (int r1, int c1) = ParseA1(topLeftA1);
            (int r2, int c2) = ParseA1(bottomRightA1);

            var relId = drawingsPart.GetIdOfPart(chartPart);

            var anchor = new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(new Xdr.ColumnId((c1 - 1).ToString()), new Xdr.ColumnOffset("0"),
                                   new Xdr.RowId((r1 - 1).ToString()), new Xdr.RowOffset("0")),
                new Xdr.ToMarker(new Xdr.ColumnId((c2 - 1).ToString()), new Xdr.ColumnOffset("0"),
                                   new Xdr.RowId((r2 - 1).ToString()), new Xdr.RowOffset("0")),
                new Xdr.GraphicFrame(
                    new Xdr.NonVisualGraphicFrameProperties(
                        new Xdr.NonVisualDrawingProperties { Id = NextDrawingId(drawingsPart), Name = "Chart " + Guid.NewGuid() },
                        new Xdr.NonVisualGraphicFrameDrawingProperties()
                    ),
                    new Xdr.Transform(new A.Offset() { X = 0, Y = 0 }, new A.Extents() { Cx = 0, Cy = 0 }),
                    new A.Graphic(new A.GraphicData(new C.ChartReference() { Id = relId })
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" })
                ),
                new Xdr.ClientData()
            );

            drawingsPart.WorksheetDrawing.Append(anchor);
            drawingsPart.WorksheetDrawing.Save();

            // 4) Worksheet'e <drawing r:id="..."> bağla (tek olsun, id doğru olsun)
            var drawingRelId = wsPart.GetIdOfPart(drawingsPart);
            var drawingEl = wsPart.Worksheet.Elements<Drawing>().FirstOrDefault();

            if (drawingEl == null)
            {
                drawingEl = new Drawing() { Id = drawingRelId };
                var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
                if (sheetData != null) wsPart.Worksheet.InsertAfter(drawingEl, sheetData);
                else wsPart.Worksheet.Append(drawingEl);
            }
            else
            {
                // varsa, r:id doğru parça mı? değilse güncelle
                if (drawingEl.Id != drawingRelId)
                    drawingEl.Id = drawingRelId;
            }

            wsPart.Worksheet.Save();
        }

        // A1 -> (row, col)
        private static (int row, int col) ParseA1(string a1)
        {
            if (string.IsNullOrWhiteSpace(a1)) return (1, 1);
            a1 = a1.Trim();
            int i = 0, col = 0;
            while (i < a1.Length && char.IsLetter(a1[i]))
            {
                col = col * 26 + (char.ToUpperInvariant(a1[i]) - 'A' + 1);
                i++;
            }
            int row = 1;
            if (i < a1.Length && int.TryParse(a1.Substring(i), out int r)) row = r;
            return (row, Math.Max(1, col));
        }

        // Tüm torunlarda non-visual id'leri tarayıp benzersiz Id üret
        private static UInt32Value NextDrawingId(DrawingsPart dp)
        {
            uint max = 0;
            if (dp.WorksheetDrawing != null)
            {
                foreach (var nvd in dp.WorksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>())
                {
                    var id = nvd.Id?.Value;
                    if (id.HasValue && id.Value > max) max = id.Value;
                }
            }
            return max + 1;
        }
    }
}

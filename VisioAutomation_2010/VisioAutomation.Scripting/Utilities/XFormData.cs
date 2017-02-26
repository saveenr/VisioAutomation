using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Drawing;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Scripting.Utilities
{
    internal struct XFormData
    {
        public double PinX;
        public double PinY;
        public double LocPinX;
        public double LocPinY;
        public double Width;
        public double Height;

        private static ColumnCell ColPinX;
        private static ColumnCell ColPinY;
        private static ColumnCell ColLocPinX;
        private static ColumnCell ColLocPinY;
        private static ColumnCell ColWidth;
        private static ColumnCell ColHeight;
        private static ShapeSheetQuery query;

        public static List<XFormData> Get(Microsoft.Office.Interop.Visio.Page page, TargetShapeIDs target)
        {
            if (query == null)
            {
                query = new ShapeSheetQuery();
                ColPinX = query.AddCell(VisioAutomation.ShapeSheet.SRCConstants.PinX, "PinX");
                ColPinY = query.AddCell(VisioAutomation.ShapeSheet.SRCConstants.PinY, "PinY");
                ColLocPinX = query.AddCell(VisioAutomation.ShapeSheet.SRCConstants.LocPinX, "LocPinX");
                ColLocPinY = query.AddCell(VisioAutomation.ShapeSheet.SRCConstants.LocPinY, "LocPinY");
                ColWidth = query.AddCell(VisioAutomation.ShapeSheet.SRCConstants.Width, "Width");
                ColHeight = query.AddCell(VisioAutomation.ShapeSheet.SRCConstants.Height, "Height");
            }

            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(page);
            var results = query.GetResults<double>(surface, target.ShapeIDs);
            if (results.Count != target.ShapeIDs.Count)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException("Didn't get as many rows back as expected");
            }
            var list = new List<XFormData>(target.ShapeIDs.Count);
            foreach (var row in results)
            {
                var xform = new XFormData();
                xform.PinX = row.Cells[ColPinX];
                xform.PinY = row.Cells[ColPinY];
                xform.LocPinX = row.Cells[ColLocPinX];
                xform.LocPinY = row.Cells[ColLocPinY];
                xform.Width = row.Cells[ColWidth];
                xform.Height = row.Cells[ColHeight];
                list.Add(xform);
            }
            return list;
        }

        public Drawing.Rectangle GetRectangle()
        {
            var pin = new Drawing.Point(this.PinX, this.PinY);
            var locpin = new Drawing.Point(this.LocPinX, this.LocPinY);
            var size = new Drawing.Size(this.Width, this.Height);
            return new Drawing.Rectangle(pin - locpin, size);
        }

        public void SetFormulas(VisioAutomation.ShapeSheet.ShapeSheetWriter writer, short id)
        {
            writer.SetFormula(id, VisioAutomation.ShapeSheet.SRCConstants.PinX, this.PinX);
            writer.SetFormula(id, VisioAutomation.ShapeSheet.SRCConstants.PinY, this.PinY);
            writer.SetFormula(id, VisioAutomation.ShapeSheet.SRCConstants.LocPinX, this.LocPinX);
            writer.SetFormula(id, VisioAutomation.ShapeSheet.SRCConstants.LocPinY, this.LocPinY);
            writer.SetFormula(id, VisioAutomation.ShapeSheet.SRCConstants.Width, this.Width);
            writer.SetFormula(id, VisioAutomation.ShapeSheet.SRCConstants.Height, this.Height);
        }

        public static Drawing.Rectangle GetBoundingBox(IEnumerable<XFormData> xfrms)
        {
            var bb = BoundingBoxBuilder.FromRectangles(xfrms.Select(x => x.GetRectangle()));
            if (!bb.HasValue)
            {
                throw new System.ArgumentException("Could not calculate bounding box");
            }
            return bb.Value;
        }
    }
}
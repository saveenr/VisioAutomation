using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Drawing;
using VisioAutomation.ShapeSheet.Query;
using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomation.Scripting.Models
{
    internal struct ShapeXFormData
    {
        public double PinX;
        public double PinY;
        public double LocPinX;
        public double LocPinY;
        public double Width;
        public double Height;

        private static CellColumn ColPinX;
        private static CellColumn ColPinY;
        private static CellColumn ColLocPinX;
        private static CellColumn ColLocPinY;
        private static CellColumn ColWidth;
        private static CellColumn ColHeight;
        private static ShapeSheetQuery query;

        public static List<ShapeXFormData> Get(Microsoft.Office.Interop.Visio.Page page, TargetShapeIDs target)
        {
            if (query == null)
            {
                query = new ShapeSheetQuery();
                ColPinX = query.AddCell(VisioAutomation.ShapeSheet.SrcConstants.XFormPinX, nameof(VisioAutomation.ShapeSheet.SrcConstants.XFormPinX));
                ColPinY = query.AddCell(VisioAutomation.ShapeSheet.SrcConstants.XFormPinY, nameof(VisioAutomation.ShapeSheet.SrcConstants.XFormPinY));
                ColLocPinX = query.AddCell(VisioAutomation.ShapeSheet.SrcConstants.XFormLocPinX, nameof(VisioAutomation.ShapeSheet.SrcConstants.XFormLocPinX));
                ColLocPinY = query.AddCell(VisioAutomation.ShapeSheet.SrcConstants.XFormLocPinY, nameof(VisioAutomation.ShapeSheet.SrcConstants.XFormLocPinY));
                ColWidth = query.AddCell(VisioAutomation.ShapeSheet.SrcConstants.XFormWidth, nameof(VisioAutomation.ShapeSheet.SrcConstants.XFormWidth));
                ColHeight = query.AddCell(VisioAutomation.ShapeSheet.SrcConstants.XFormHeight, nameof(VisioAutomation.ShapeSheet.SrcConstants.XFormHeight));
            }

            var results = query.GetResults<double>(page, target.ShapeIDs);
            if (results.Count != target.ShapeIDs.Count)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException("Didn't get as many rows back as expected");
            }
            var list = new List<ShapeXFormData>(target.ShapeIDs.Count);
            foreach (var row in results)
            {
                var xform = new ShapeXFormData();
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

        public void SetFormulas(SidSrcWriter writer, short id)
        {
            writer.SetFormula(id, VisioAutomation.ShapeSheet.SrcConstants.XFormPinX, this.PinX);
            writer.SetFormula(id, VisioAutomation.ShapeSheet.SrcConstants.XFormPinY, this.PinY);
            writer.SetFormula(id, VisioAutomation.ShapeSheet.SrcConstants.XFormLocPinX, this.LocPinX);
            writer.SetFormula(id, VisioAutomation.ShapeSheet.SrcConstants.XFormLocPinY, this.LocPinY);
            writer.SetFormula(id, VisioAutomation.ShapeSheet.SrcConstants.XFormWidth, this.Width);
            writer.SetFormula(id, VisioAutomation.ShapeSheet.SrcConstants.XFormHeight, this.Height);
        }

        public static Drawing.Rectangle GetBoundingBox(IEnumerable<ShapeXFormData> xfrms)
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
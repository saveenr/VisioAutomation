using System.Collections.Generic;
using System.Linq;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Models
{
    internal struct ShapeXFormData
    {
        public double XFormPinX;
        public double XFormPinY;
        public double XFormLocPinX;
        public double XFormLocPinY;
        public double XFormWidth;
        public double XFormHeight;

        private static VASS.Query.Column ColXFormPinX;
        private static VASS.Query.Column ColXFormPinY;
        private static VASS.Query.Column ColXFormLocPinX;
        private static VASS.Query.Column ColXFormLocPinY;
        private static VASS.Query.Column ColXFormWidth;
        private static VASS.Query.Column ColXFormHeight;
        private static VASS.Query.CellQuery query;

        public static List<ShapeXFormData> Get(IVisio.Page page, TargetShapeIDs target)
        {
            if (query == null)
            {
                query = new VASS.Query.CellQuery();
                ColXFormPinX = query.Columns.Add(VASS.SrcConstants.XFormPinX, nameof(ShapeXFormData.XFormPinX));
                ColXFormPinY = query.Columns.Add(VASS.SrcConstants.XFormPinY, nameof(ShapeXFormData.XFormPinY));
                ColXFormLocPinX = query.Columns.Add(VASS.SrcConstants.XFormLocPinX, nameof(ShapeXFormData.XFormLocPinX));
                ColXFormLocPinY = query.Columns.Add(VASS.SrcConstants.XFormLocPinY, nameof(ShapeXFormData.XFormLocPinY));
                ColXFormWidth = query.Columns.Add(VASS.SrcConstants.XFormWidth, nameof(ShapeXFormData.XFormWidth));
                ColXFormHeight = query.Columns.Add(VASS.SrcConstants.XFormHeight, nameof(ShapeXFormData.XFormHeight));
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
                xform.XFormPinX = row[ColXFormPinX];
                xform.XFormPinY = row[ColXFormPinY];
                xform.XFormLocPinX = row[ColXFormLocPinX];
                xform.XFormLocPinY = row[ColXFormLocPinY];
                xform.XFormWidth = row[ColXFormWidth];
                xform.XFormHeight = row[ColXFormHeight];
                list.Add(xform);
            }
            return list;
        }

        public VisioAutomation.Geometry.Rectangle GetRectangle()
        {
            var pin = new VisioAutomation.Geometry.Point(this.XFormPinX, this.XFormPinY);
            var locpin = new VisioAutomation.Geometry.Point(this.XFormLocPinX, this.XFormLocPinY);
            var size = new VisioAutomation.Geometry.Size(this.XFormWidth, this.XFormHeight);
            return new VisioAutomation.Geometry.Rectangle(pin - locpin, size);
        }

        public void SetFormulas(VASS.Writers.SidSrcWriter writer, short id)
        {
            writer.SetFormula(id, VASS.SrcConstants.XFormPinX, this.XFormPinX);
            writer.SetFormula(id, VASS.SrcConstants.XFormPinY, this.XFormPinY);
            writer.SetFormula(id, VASS.SrcConstants.XFormLocPinX, this.XFormLocPinX);
            writer.SetFormula(id, VASS.SrcConstants.XFormLocPinY, this.XFormLocPinY);
            writer.SetFormula(id, VASS.SrcConstants.XFormWidth, this.XFormWidth);
            writer.SetFormula(id, VASS.SrcConstants.XFormHeight, this.XFormHeight);
        }

        public static VisioAutomation.Geometry.Rectangle GetBoundingBox(IEnumerable<ShapeXFormData> xfrms)
        {
            var bb = VisioAutomation.Models.Geometry.BoundingBoxBuilder.FromRectangles(xfrms.Select(x => x.GetRectangle()));
            if (!bb.HasValue)
            {
                throw new System.ArgumentException("Could not calculate bounding box");
            }
            return bb.Value;
        }
    }
}
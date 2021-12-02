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

        private static VASS.Data.Column _static_col_x_form_pin_x;
        private static VASS.Data.Column _static_col_x_form_pin_y;
        private static VASS.Data.Column _static_col_x_form_loc_pin_x;
        private static VASS.Data.Column _static_col_x_form_loc_pin_y;
        private static VASS.Data.Column _static_col_x_form_width;
        private static VASS.Data.Column _static_col_x_form_height;
        private static VASS.Query.CellQuery _static_query;

        internal static List<ShapeXFormData> _get_xfrms(IVisio.Page page, IList<int> shapeids)
        {
            if (_static_query == null)
            {
                _static_query = new VASS.Query.CellQuery();
                var cols = _static_query.Columns;
                _static_col_x_form_pin_x = cols.Add(VisioAutomation.Core.SrcConstants.XFormPinX);
                _static_col_x_form_pin_y = cols.Add(VisioAutomation.Core.SrcConstants.XFormPinY);
                _static_col_x_form_loc_pin_x = cols.Add(VisioAutomation.Core.SrcConstants.XFormLocPinX);
                _static_col_x_form_loc_pin_y = cols.Add(VisioAutomation.Core.SrcConstants.XFormLocPinY);
                _static_col_x_form_width = cols.Add(VisioAutomation.Core.SrcConstants.XFormWidth);
                _static_col_x_form_height = cols.Add(VisioAutomation.Core.SrcConstants.XFormHeight);
            }

            var results = _static_query.GetResults<double>(page, shapeids);
            if (results.Count != shapeids.Count)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException("Didn't get as many rows back as expected");
            }
            var list = new List<ShapeXFormData>(shapeids.Count);
            foreach (var row in results)
            {
                var xform = new ShapeXFormData();
                xform.XFormPinX = row[_static_col_x_form_pin_x];
                xform.XFormPinY = row[_static_col_x_form_pin_y];
                xform.XFormLocPinX = row[_static_col_x_form_loc_pin_x];
                xform.XFormLocPinY = row[_static_col_x_form_loc_pin_y];
                xform.XFormWidth = row[_static_col_x_form_width];
                xform.XFormHeight = row[_static_col_x_form_height];
                list.Add(xform);
            }
            return list;
        }

        public VisioAutomation.Core.Rectangle GetRectangle()
        {
            var pin = new VisioAutomation.Core.Point(this.XFormPinX, this.XFormPinY);
            var locpin = new VisioAutomation.Core.Point(this.XFormLocPinX, this.XFormLocPinY);
            var size = new VisioAutomation.Core.Size(this.XFormWidth, this.XFormHeight);
            return new VisioAutomation.Core.Rectangle(pin - locpin, size);
        }

        public void SetFormulas(VASS.Writers.SidSrcWriter writer, short id)
        {
            writer.SetValue(id, VisioAutomation.Core.SrcConstants.XFormPinX, this.XFormPinX);
            writer.SetValue(id, VisioAutomation.Core.SrcConstants.XFormPinY, this.XFormPinY);
            writer.SetValue(id, VisioAutomation.Core.SrcConstants.XFormLocPinX, this.XFormLocPinX);
            writer.SetValue(id, VisioAutomation.Core.SrcConstants.XFormLocPinY, this.XFormLocPinY);
            writer.SetValue(id, VisioAutomation.Core.SrcConstants.XFormWidth, this.XFormWidth);
            writer.SetValue(id, VisioAutomation.Core.SrcConstants.XFormHeight, this.XFormHeight);
        }

        public static VisioAutomation.Core.Rectangle GetBoundingBox(IEnumerable<ShapeXFormData> xfrms)
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
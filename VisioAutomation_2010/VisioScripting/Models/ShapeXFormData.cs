
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

        private static VASS.Query.Column _static_col_x_form_pin_x;
        private static VASS.Query.Column _static_col_x_form_pin_y;
        private static VASS.Query.Column _static_col_x_form_loc_pin_x;
        private static VASS.Query.Column _static_col_x_form_loc_pin_y;
        private static VASS.Query.Column _static_col_x_form_width;
        private static VASS.Query.Column _static_col_x_form_height;
        private static VASS.Query.CellQuery _static_query;

        internal static List<ShapeXFormData> _get_xfrms(IVisio.Page page, IList<int> shapeids)
        {
            if (_static_query == null)
            {
                _static_query = new VASS.Query.CellQuery();
                var cols = _static_query.Columns;
                _static_col_x_form_pin_x = cols.Add(VASS.SrcConstants.XFormPinX, nameof(ShapeXFormData.XFormPinX));
                _static_col_x_form_pin_y = cols.Add(VASS.SrcConstants.XFormPinY, nameof(ShapeXFormData.XFormPinY));
                _static_col_x_form_loc_pin_x = cols.Add(VASS.SrcConstants.XFormLocPinX, nameof(ShapeXFormData.XFormLocPinX));
                _static_col_x_form_loc_pin_y = cols.Add(VASS.SrcConstants.XFormLocPinY, nameof(ShapeXFormData.XFormLocPinY));
                _static_col_x_form_width = cols.Add(VASS.SrcConstants.XFormWidth, nameof(ShapeXFormData.XFormWidth));
                _static_col_x_form_height = cols.Add(VASS.SrcConstants.XFormHeight, nameof(ShapeXFormData.XFormHeight));
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

        public VisioAutomation.Geometry.Rectangle GetRectangle()
        {
            var pin = new VisioAutomation.Geometry.Point(this.XFormPinX, this.XFormPinY);
            var locpin = new VisioAutomation.Geometry.Point(this.XFormLocPinX, this.XFormLocPinY);
            var size = new VisioAutomation.Geometry.Size(this.XFormWidth, this.XFormHeight);
            return new VisioAutomation.Geometry.Rectangle(pin - locpin, size);
        }

        public void SetFormulas(VASS.Writers.SidSrcWriter writer, short id)
        {
            writer.SetValue(id, VASS.SrcConstants.XFormPinX, this.XFormPinX);
            writer.SetValue(id, VASS.SrcConstants.XFormPinY, this.XFormPinY);
            writer.SetValue(id, VASS.SrcConstants.XFormLocPinX, this.XFormLocPinX);
            writer.SetValue(id, VASS.SrcConstants.XFormLocPinY, this.XFormLocPinY);
            writer.SetValue(id, VASS.SrcConstants.XFormWidth, this.XFormWidth);
            writer.SetValue(id, VASS.SrcConstants.XFormHeight, this.XFormHeight);
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
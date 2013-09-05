using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout
{
    public static class LayoutHelper
    {
        private static VA.Drawing.Rectangle GetRectangle(XFormCells xFormCells)
        {
            var pin = new VA.Drawing.Point(xFormCells.PinX.Result, xFormCells.PinY.Result);
            var locpin = new VA.Drawing.Point(xFormCells.LocPinX.Result, xFormCells.LocPinY.Result);
            var size = new VA.Drawing.Size(xFormCells.Width.Result, xFormCells.Height.Result);
            return new VA.Drawing.Rectangle(pin - locpin, size);
        }

        private static double GetPosition(VA.Layout.XFormCells xform, XFormPosition pos)
        {
            if (pos == XFormPosition.PinY)
            {
                return xform.PinY.Result;
            }
            else if (pos == XFormPosition.PinX)
            {
                return xform.PinX.Result;
            }
            else
            {
                var r = GetRectangle(xform);
                if (pos == XFormPosition.Left)
                {
                    return r.Left;
                }
                else if (pos == XFormPosition.Right)
                {
                    return r.Right;
                }
                else if (pos == XFormPosition.Top)
                {
                    return r.Top;
                }
                else if (pos == XFormPosition.Right)
                {
                    return r.Bottom;
                }
                else
                {
                    throw new System.ArgumentOutOfRangeException("pos");
                }
            }
        }

        public static IList<int> OrderShapesByXFormPosition(IVisio.Page page, IList<int> shapeids, XFormPosition pos)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            if (shapeids == null)
            {
                throw new System.ArgumentNullException("shapeids");
            }

            // First get the transforms of the shapes on the given axis
            var xforms = VA.Layout.XFormCells.GetCells(page, shapeids);

            // Then, sort the shapeids pased on the corresponding value in the results


            var sorted_shape_ids = Enumerable.Range(0, shapeids.Count)
                .Select(i => new {index = i, shapeid = shapeids[i], pos = GetPosition(xforms[i], pos)})
                .OrderBy(i => i.pos)
                .Select(i=>i.shapeid)
                .ToList();

            return sorted_shape_ids;
        }

        public static void DistributeWithSpacing(IVisio.Page page, IList<int> shapeids, VA.Drawing.Axis axis, double spacing)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            if (shapeids == null)
            {
                throw new System.ArgumentNullException("shapeids");
            }

            if (spacing < 0.0)
            {
                throw new System.ArgumentOutOfRangeException("spacing");
            }

            if (shapeids.Count < 2)
            {
                return;
            }

            // Calculate the new Xfrms
            var sortpos = axis == VA.Drawing.Axis.XAxis
                              ? VA.Layout.XFormPosition.PinX
                              : VA.Layout.XFormPosition.PinY;

            var delta = axis == VA.Drawing.Axis.XAxis
                            ? new VA.Drawing.Size(spacing, 0)
                            : new VA.Drawing.Size(0, spacing);


            var sorted_shape_ids = VA.Layout.LayoutHelper.OrderShapesByXFormPosition(page, shapeids, sortpos);
            var input_xfrms = VA.Layout.XFormCells.GetCells(page, sorted_shape_ids);
            var output_xfrms = new List<VA.Layout.XFormCells>(input_xfrms.Count);
            var bb = GetBoundingBox(input_xfrms);
            var cur_pos = new VA.Drawing.Point(bb.Left, bb.Bottom);

            foreach (var input_xfrm in input_xfrms)
            {
                var new_pinpos = axis == VA.Drawing.Axis.XAxis
                                     ? new VA.Drawing.Point(cur_pos.X + input_xfrm.LocPinX.Result, input_xfrm.PinY.Result)
                                     : new VA.Drawing.Point(input_xfrm.PinX.Result, cur_pos.Y + input_xfrm.LocPinY.Result);

                var output_xfrm = new VA.Layout.XFormCells();
                output_xfrm.PinX = new_pinpos.X;
                output_xfrm.PinY = new_pinpos.Y;
                output_xfrms.Add(output_xfrm);

                cur_pos = cur_pos.Add(input_xfrm.Width.Result, input_xfrm.Height.Result).Add(delta);
            }

            // Apply the changes
            update_xfrms(page,sorted_shape_ids,output_xfrms);
        }

        private static void update_xfrms(IVisio.Page page, IList<int> shapeids, IList<VA.Layout.XFormCells> xfrms)
        {
            
            var update = new VA.ShapeSheet.Update();
            for (int i = 0; i < shapeids.Count; i++)
            {
                var shape_id = shapeids[i];
                var xfrm = xfrms[i];
                update.SetFormulas((short)shape_id, xfrm);
            }
            update.Execute(page);
        }

        public static VA.Drawing.Rectangle GetBoundingBox(IEnumerable<VA.Layout.XFormCells> xfrms)
        {
            var bb = new VA.Drawing.BoundingBox(xfrms.Select(VA.Layout.LayoutHelper.GetRectangle));
            if (!bb.HasValue)
            {
                throw new System.ArgumentException("Could not calculate bounding box");
            }
            else
            {
                return bb.Rectangle;
            }
        }

        public static void SnapCorner(IVisio.Page page, IList<int> shapeids, VA.Drawing.Size snapsize, SnapCornerPosition corner)
        {
            // First caculate the new transforms
            var snap_grid = new VA.Drawing.SnappingGrid(snapsize);
            var input_xfrms = VA.Layout.XFormCells.GetCells(page, shapeids);
            var output_xfrms = new List<VA.Layout.XFormCells>(input_xfrms.Count);

            foreach (var input_xfrm in input_xfrms)
            {
                var old_lower_left = VA.Layout.LayoutHelper.GetRectangle(input_xfrm).LowerLeft;
                var new_lower_left = snap_grid.Snap(old_lower_left);
                var output_xfrm = _SnapCorner(corner, new_lower_left, input_xfrm);
                output_xfrms.Add(output_xfrm);
            }

            // Now apply them
            update_xfrms(page, shapeids, output_xfrms);
        }

        private static XFormCells _SnapCorner(SnapCornerPosition corner, VA.Drawing.Point new_lower_left, XFormCells input_xfrm)
        {
            var new_pin_position = GetPinPositionForCorner(input_xfrm, new_lower_left, corner);

            var output_xfrm = new VA.Layout.XFormCells();
            if (new_pin_position.X != input_xfrm.PinX.Result)
            {
                output_xfrm.PinX = new_pin_position.X;
            }
            if (new_pin_position.Y != input_xfrm.PinY.Result)
            {
                output_xfrm.PinY = new_pin_position.Y;
            }
            return output_xfrm;
        }

        private static VA.Drawing.Point GetPinPositionForCorner( VA.Layout.XFormCells input_xfrm, VA.Drawing.Point new_lower_left, SnapCornerPosition corner)
        {
            var size = new VA.Drawing.Size(input_xfrm.Width.Result, input_xfrm.Height.Result);
            var locpin = new VA.Drawing.Point(input_xfrm.LocPinX.Result, input_xfrm.LocPinY.Result);

            switch (corner)
            {
                case SnapCornerPosition.LowerLeft:
                    {
                        return new_lower_left.Add(locpin.X, locpin.Y);
                    }
                case SnapCornerPosition.UpperRight:
                    {
                        return new_lower_left.Subtract(size.Width, size.Height).Add(locpin.X, locpin.Y);
                    }
                case SnapCornerPosition.LowerRight:
                    {
                        return new_lower_left.Subtract(size.Width, 0).Add(locpin.X, locpin.Y);
                    }
                case SnapCornerPosition.UpperLeft:
                    {
                        return new_lower_left.Subtract(0, size.Height).Add(locpin.X, locpin.Y);
                    }
                default:
                    {
                        throw new System.ArgumentOutOfRangeException("corner", "Unsupported corner");
                    }
            }
        }

        public static void SnapSize(IVisio.Page page, IList<int> shapeids, VA.Drawing.Size snapsize, VA.Drawing.Size minsize)
        {
            var input_xfrms = VA.Layout.XFormCells.GetCells(page, shapeids);
            var output_xfrms = new List<VA.Layout.XFormCells>(input_xfrms.Count);

            var grid = new VA.Drawing.SnappingGrid(snapsize);
            foreach (var input_xfrm in input_xfrms)
            {
                var inut_size = new VA.Drawing.Size( input_xfrm.Width.Result, input_xfrm.Height.Result );
                var snapped_size = grid.Snap(inut_size);
                double max_w = System.Math.Max(snapped_size.Width, minsize.Width);
                double max_h = System.Math.Max(snapped_size.Height, minsize.Height);
                var new_size = new VA.Drawing.Size(max_w, max_h);

                var output_xfrm = new VA.Layout.XFormCells();
                output_xfrm.Width = new_size.Width;
                output_xfrm.Height = new_size.Height;

                output_xfrms.Add(output_xfrm);
            }

            // Now apply them
            update_xfrms(page, shapeids, output_xfrms);
        }
    }
}
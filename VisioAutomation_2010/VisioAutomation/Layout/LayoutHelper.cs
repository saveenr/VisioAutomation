using System;
using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout
{
    public static class LayoutHelper
    {
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
            else if (pos == XFormPosition.Left)
            {
                return xform.Rect.Left;
            }
            else if (pos == XFormPosition.Right)
            {
                return xform.Rect.Right;
            }
            else if (pos == XFormPosition.Top)
            {
                return xform.Rect.Top;
            }
            else if (pos == XFormPosition.Right)
            {
                return xform.Rect.Bottom;
            }
            else
            {
                throw new System.ArgumentOutOfRangeException("pos");
            }
        }

        public static IList<int> SortShapesByPosition(IVisio.Page page,
                                                 IList<int> shapeids,
                                                 XFormPosition pos)
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
            var xforms = VA.Layout.LayoutHelper.GetXForm(page, shapeids);

            // Then, sort the shapeids pased on the corresponding value in the results


            var sorted_shape_ids = Enumerable.Range(0, shapeids.Count)
                .Select(i => new {index = i, shapeid = shapeids[i], pos = GetPosition(xforms[i], pos)})
                .OrderBy(i => i.pos)
                .Select(i=>i.shapeid)
                .ToList();

            return sorted_shape_ids;
        }

        public static void DistributeWithSpacing(IVisio.Page page,
                                     IList<int> shapeids,
                                     VA.Drawing.Axis axis,
                                     double spacing)
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

            var sortpos = axis == VA.Drawing.Axis.XAxis
                              ? VA.Layout.XFormPosition.PinX
                              : VA.Layout.XFormPosition.PinY;

            var delta = axis == VA.Drawing.Axis.XAxis
                            ? new VA.Drawing.Size(spacing, 0)
                            : new VA.Drawing.Size(0, spacing);


            var sorted_shape_ids = VA.Layout.LayoutHelper.SortShapesByPosition(page, shapeids, sortpos);
            var xfrms = VA.Layout.LayoutHelper.GetXForm(page, sorted_shape_ids); ;
            var bb = GetBoundingBox(xfrms);
            var cur_pos = new VA.Drawing.Point(bb.Left, bb.Bottom);

            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            for (int i = 0; i < sorted_shape_ids.Count; i++)
            {
                var shape_id = sorted_shape_ids[i];
                var xfrm = xfrms[i];

                var new_pinpos = axis == VA.Drawing.Axis.XAxis
                                     ? new VA.Drawing.Point(cur_pos.X + xfrm.LocPinX.Result, xfrm.PinY.Result)
                                     : new VA.Drawing.Point(xfrm.PinX.Result, cur_pos.Y + xfrm.LocPinY.Result);

                update.SetFormula((short)shape_id, VA.ShapeSheet.SRCConstants.PinX, new_pinpos.X);
                update.SetFormula((short)shape_id, VA.ShapeSheet.SRCConstants.PinY, new_pinpos.Y);
                cur_pos = cur_pos.Add(xfrm.Width.Result,xfrm.Height.Result).Add(delta);
            }

            update.Execute(page);
        }

        public static VA.Drawing.Rectangle GetBoundingBox(IEnumerable<VA.Layout.XFormCells> xfrms)
        {
            var bb = new VA.Drawing.BoundingBox(xfrms.Select(i => i.Rect));
            if (!bb.HasValue)
            {
                throw new System.ArgumentException("Could not calculate bounding box");
            }
            else
            {
                return bb.Rectangle;
            }
        }

        public static void SnapCorner(IVisio.Page page,
                              IList<int> shapeids,
                              VA.Drawing.Size snapsize,
                              SnapCornerPosition corner)
        {
            var layout_info = VA.Layout.LayoutHelper.GetXForm(page, shapeids);
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            var snap_grid = new VA.Drawing.SnappingGrid(snapsize);

            foreach (int i in Enumerable.Range(0, shapeids.Count))
            {
                var shapeid = shapeids[i];
                var old_layout = layout_info[i];
                var old_bb = old_layout.Rect;
                var old_bb_pos = old_bb.LowerLeft;

                var new_corner_pos = snap_grid.Snap(old_bb_pos);

                var new_pin_position = GetPinPositionForCorner(
                    new VA.Drawing.Point( old_layout.PinX.Result, old_layout.PinY.Result ),
                    old_layout.Size,
                    old_layout.LocPin,
                    new_corner_pos,
                    corner);

                if (new_pin_position.X != old_layout.PinX.Result)
                {
                    update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.PinX, new_pin_position.X);
                }

                if (new_pin_position.Y != old_layout.PinY.Result)
                {
                    update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.PinY, new_pin_position.Y);
                }
            }

            update.Execute(page);
        }

        private static VA.Drawing.Point GetPinPositionForCorner(
            VA.Drawing.Point pinpos,
            VA.Drawing.Size size,
            VA.Drawing.Point locpin,
            VA.Drawing.Point new_corner_pos,
            SnapCornerPosition corner)
        {
            switch (corner)
            {
                case SnapCornerPosition.LowerLeft:
                    {
                        return new_corner_pos.Add(locpin.X, locpin.Y);
                    }
                case SnapCornerPosition.UpperRight:
                    {
                        return new_corner_pos.Subtract(size.Width, size.Height).Add(locpin.X, locpin.Y);
                    }
                case SnapCornerPosition.LowerRight:
                    {
                        return new_corner_pos.Subtract(size.Width, 0).Add(locpin.X, locpin.Y);
                    }
                case SnapCornerPosition.UpperLeft:
                    {
                        return new_corner_pos.Subtract(0, size.Height).Add(locpin.X, locpin.Y);
                    }
                default:
                    {
                        throw new System.ArgumentOutOfRangeException("corner", "Unsupported corner");
                    }
            }
        }

        public static void SnapSize(IVisio.Page page, IList<int> shapeids, VA.Drawing.Size snapsize, VA.Drawing.Size minsize)
        {
            var layout_info = VA.Layout.LayoutHelper.GetXForm(page, shapeids);
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            var grid = new VA.Drawing.SnappingGrid(snapsize);

            for (int i = 0; i < shapeids.Count; i++)
            {
                int shapeid = shapeids[i];
                var old_size = layout_info[i].Size;
                var snapped_size = grid.Snap(old_size);
                double max_w = System.Math.Max(snapped_size.Width, minsize.Width);
                double max_h = System.Math.Max(snapped_size.Height, minsize.Height);
                var new_size = new VA.Drawing.Size(max_w, max_h);

                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.Width, new_size.Width);
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.Height, new_size.Height);
            }

            update.Execute(page);
        }

        public static void AlignTo(IVisio.Page page, IList<int> shapeids, VA.Drawing.AlignmentHorizontal align, double x)
        {
            var xfrms = VA.Layout.LayoutHelper.GetXForm(page, shapeids);
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            for (int i = 0; i < shapeids.Count; i++)
            {
                var info = xfrms[i];
                double nx = 0.0;
                if (align == VA.Drawing.AlignmentHorizontal.Left)
                {
                    nx = x + info.LocPinX.Result;
                }
                else if (align == VA.Drawing.AlignmentHorizontal.Center)
                {
                    nx = x + info.LocPinX.Result - (info.Width.Result / 2.0);
                }
                else if (align == VA.Drawing.AlignmentHorizontal.Right)
                {
                    nx = x + info.LocPinX.Result - info.Width.Result;
                }
                update.SetFormula((short)shapeids[i], VA.ShapeSheet.SRCConstants.PinX, nx);

            }

            update.Execute(page);
        }

        public static void AlignTo(IVisio.Page page, IList<int> shapeids, VA.Drawing.AlignmentVertical align, double y)
        {
            var xfrms = VA.Layout.LayoutHelper.GetXForm(page, shapeids);
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            for (int i = 0; i < shapeids.Count; i++)
            {
                var info = xfrms[i];
                double ny = 0.0;

                if (align == VA.Drawing.AlignmentVertical.Top)
                {
                    ny = y + info.LocPinY.Result - info.Height.Result;
                }
                else if (align == VA.Drawing.AlignmentVertical.Center)
                {
                    ny = y + info.LocPinY.Result - (info.Height.Result / 2.0);
                }
                else if (align == VA.Drawing.AlignmentVertical.Bottom)
                {
                    ny = y + info.LocPinY.Result;
                }

                update.SetFormula((short)shapeids[i], VA.ShapeSheet.SRCConstants.PinY, ny);
            }

            update.Execute(page);
        }

        public static VA.Layout.XFormCells GetXForm(IVisio.Shape shape)
        {
            return XFormCells.GetCells(shape);
        }

        public static IList<VA.Layout.XFormCells> GetXForm(IVisio.Page page, IList<int> shapeids)
        {
            return XFormCells.GetCells(page, shapeids);
        }

        public static VA.Layout.LockCells GetLock(IVisio.Shape shape)
        {
            return LockCells.GetCells(shape);
        }

        public static IList<VA.Layout.ShapeLayoutCells> GetShapeLayout(IVisio.Page page, IList<int> shapeids)
        {
            return ShapeLayoutCells.GetCells(page, shapeids);
        }

        public static VA.Layout.ShapeLayoutCells GetShapeLayout(IVisio.Shape shape)
        {
            return ShapeLayoutCells.GetCells(shape);
        }

        public static IList<VA.Layout.LockCells> GetLock(IVisio.Page page, IList<int> shapeids)
        {
            return LockCells.GetCells(page, shapeids);
        }

        public static void SendShapes( IVisio.Selection selection, VA.Layout.ShapeSendDirection dir)
        {

            if (dir == VA.Layout.ShapeSendDirection.ToBack)
            {
                selection.SendToBack();
            }
            else if (dir == VA.Layout.ShapeSendDirection.Backward)
            {
                selection.SendBackward();
            }
            else if (dir == VA.Layout.ShapeSendDirection.Forward)
            {
                selection.BringForward();
            }
            else if (dir == VA.Layout.ShapeSendDirection.ToFront)
            {
                selection.BringToFront();
            }
        }
    }
}
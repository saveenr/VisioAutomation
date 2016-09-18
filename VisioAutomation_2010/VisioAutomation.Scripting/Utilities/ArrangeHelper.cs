using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Utilities
{
    internal static class ArrangeHelper
    {
        public struct XForm
        {
            public double PinX;
            public double PinY;
            public double LocPinX;
            public double LocPinY;
            public double Width;
            public double Height;

            private static VisioAutomation.ShapeSheet.Queries.Columns.ColumnQuery ColPinX;
            private static VisioAutomation.ShapeSheet.Queries.Columns.ColumnQuery ColPinY;
            private static VisioAutomation.ShapeSheet.Queries.Columns.ColumnQuery ColLocPinX;
            private static VisioAutomation.ShapeSheet.Queries.Columns.ColumnQuery ColLocPinY;
            private static VisioAutomation.ShapeSheet.Queries.Columns.ColumnQuery ColWidth;
            private static VisioAutomation.ShapeSheet.Queries.Columns.ColumnQuery ColHeight;
            private static VisioAutomation.ShapeSheet.Queries.Query query;

            public static List<XForm> Get(TargetShapeIDs target)
            {
                if (query == null)
                {
                    query = new VisioAutomation.ShapeSheet.Queries.Query();
                    ColPinX = query.AddCell(VisioAutomation.ShapeSheet.SRCConstants.PinX, "PinX");
                    ColPinY = query.AddCell(VisioAutomation.ShapeSheet.SRCConstants.PinY, "PinY");
                    ColLocPinX = query.AddCell(VisioAutomation.ShapeSheet.SRCConstants.LocPinX, "LocPinX");
                    ColLocPinY = query.AddCell(VisioAutomation.ShapeSheet.SRCConstants.LocPinY, "LocPinY");
                    ColWidth = query.AddCell(VisioAutomation.ShapeSheet.SRCConstants.Width, "Width");
                    ColHeight = query.AddCell(VisioAutomation.ShapeSheet.SRCConstants.Height, "Height");
                }

                var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(target.Page);
                var results = query.GetResults<double>(surface, target.ShapeIDs);
                if (results.Count != target.ShapeIDs.Count)
                {
                    throw new VisioAutomation.Exceptions.InternalAssertionException("Didn't get as many rows back as expected");
                }
                var list = new List<XForm>(target.ShapeIDs.Count);
                foreach (var row in results)
                {
                    var xform = new XForm();
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

            public void SetFormulas(VisioAutomation.ShapeSheet.Writers.FormulaWriterSIDSRC writer, short id)
            {
                writer.SetFormula(id, VisioAutomation.ShapeSheet.SRCConstants.PinX, this.PinX);
                writer.SetFormula(id, VisioAutomation.ShapeSheet.SRCConstants.PinY, this.PinY);
                writer.SetFormula(id, VisioAutomation.ShapeSheet.SRCConstants.LocPinX, this.LocPinX);
                writer.SetFormula(id, VisioAutomation.ShapeSheet.SRCConstants.LocPinY, this.LocPinY);
                writer.SetFormula(id, VisioAutomation.ShapeSheet.SRCConstants.Width, this.Width);
                writer.SetFormula(id, VisioAutomation.ShapeSheet.SRCConstants.Height, this.Height);
            }

            public static Drawing.Rectangle GetBoundingBox(IEnumerable<XForm> xfrms)
            {
                var bb = new VisioAutomation.Drawing.Layout.BoundingBox(xfrms.Select(x => x.GetRectangle()));
                if (!bb.HasValue)
                {
                    throw new System.ArgumentException("Could not calculate bounding box");
                }
                return bb.Rectangle;
            }
        }

        private static double GetPositionOnShape(XForm xform, VisioAutomation.Scripting.Layout.RelativePosition pos)
        {
            var r = xform.GetRectangle();

            switch (pos)
            {
                case VisioAutomation.Scripting.Layout.RelativePosition.PinY:
                    return xform.PinY;
                case VisioAutomation.Scripting.Layout.RelativePosition.PinX:
                    return xform.PinX;
                case VisioAutomation.Scripting.Layout.RelativePosition.Left:
                    return r.Left;
                case VisioAutomation.Scripting.Layout.RelativePosition.Right:
                    return r.Right;
                case VisioAutomation.Scripting.Layout.RelativePosition.Top:
                    return r.Top;
                case VisioAutomation.Scripting.Layout.RelativePosition.Bottom:
                    return r.Bottom;
            }

            throw new System.ArgumentOutOfRangeException(nameof(pos));
        }

        internal static IList<int> SortShapesByPosition(TargetShapeIDs targets, VisioAutomation.Scripting.Layout.RelativePosition pos)
        {
            // First get the transforms of the shapes on the given axis
            var xforms = XForm.Get(targets);

            // Then, sort the shapeids pased on the corresponding value in the results

            var sorted_shape_ids = Enumerable.Range(0, targets.ShapeIDs.Count)
                .Select(i => new { index = i, shapeid = targets.ShapeIDs[i], pos = ArrangeHelper.GetPositionOnShape(xforms[i], pos) })
                .OrderBy(i => i.pos)
                .Select(i => i.shapeid)
                .ToList();

            return sorted_shape_ids;
        }

        public static void DistributeWithSpacing(TargetShapeIDs target, VisioAutomation.Drawing.Layout.Axis axis, double spacing)
        {
            if (spacing < 0.0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(spacing));
            }

            if (target.ShapeIDs.Count < 2)
            {
                return;
            }

            // Calculate the new Xfrms
            var sortpos = axis == VisioAutomation.Drawing.Layout.Axis.XAxis
                ? VisioAutomation.Scripting.Layout.RelativePosition.PinX
                : VisioAutomation.Scripting.Layout.RelativePosition.PinY;

            var delta = axis == VisioAutomation.Drawing.Layout.Axis.XAxis
                ? new Drawing.Size(spacing, 0)
                : new Drawing.Size(0, spacing);


            var input_xfrms = XForm.Get(target);
            var bb = XForm.GetBoundingBox(input_xfrms);
            var cur_pos = new Drawing.Point(bb.Left, bb.Bottom);

            var newpositions = new List<VisioAutomation.Drawing.Point>(target.ShapeIDs.Count);
            foreach (var input_xfrm in input_xfrms)
            {
                var new_pinpos = axis == VisioAutomation.Drawing.Layout.Axis.XAxis
                    ? new Drawing.Point(cur_pos.X + input_xfrm.LocPinX, input_xfrm.PinY)
                    : new Drawing.Point(input_xfrm.PinX, cur_pos.Y + input_xfrm.LocPinY);

                newpositions.Add(new_pinpos);
                cur_pos = cur_pos.Add(input_xfrm.Width, input_xfrm.Height).Add(delta);
            }

            // Apply the changes
            var sorted_shape_ids = ArrangeHelper.SortShapesByPosition(target, sortpos);

            ModifyPinPositions(target.Page, sorted_shape_ids, newpositions);
        }

        private static void ModifyPinPositions(IVisio.Page page, IList<int> sorted_shape_ids, List<VisioAutomation.Drawing.Point> newpositions)
        {
            var writer = new VisioAutomation.ShapeSheet.Writers.FormulaWriterSIDSRC();
            for (int i = 0; i < newpositions.Count; i++)
            {
                writer.SetFormula((short)sorted_shape_ids[i], VisioAutomation.ShapeSheet.SRCConstants.PinX, newpositions[i].X);
                writer.SetFormula((short)sorted_shape_ids[i], VisioAutomation.ShapeSheet.SRCConstants.PinY, newpositions[i].Y);
            }
            writer.Commit(page);
        }

        private static void ModifySizes(IVisio.Page page, IList<int> sorted_shape_ids, List<VisioAutomation.Drawing.Size> newsizes)
        {
            var writer = new VisioAutomation.ShapeSheet.Writers.FormulaWriterSIDSRC();
            for (int i = 0; i < newsizes.Count; i++)
            {
                writer.SetFormula((short)sorted_shape_ids[i], VisioAutomation.ShapeSheet.SRCConstants.Width, newsizes[i].Width);
                writer.SetFormula((short)sorted_shape_ids[i], VisioAutomation.ShapeSheet.SRCConstants.Height, newsizes[i].Height);
            }
            writer.Commit(page);
        }

        public static void SnapCorner(TargetShapeIDs target, Drawing.Size snapsize, VisioAutomation.Scripting.Layout.SnapCornerPosition corner)
        {
            // First caculate the new transforms
            var snap_grid = new SnappingGrid(snapsize);
            var input_xfrms = XForm.Get(target);
            var output_xfrms = new List<VisioAutomation.Drawing.Point>(input_xfrms.Count);

            foreach (var input_xfrm in input_xfrms)
            {
                var old_rect = input_xfrm.GetRectangle();
                var old_lower_left = old_rect.LowerLeft;
                var new_lower_left = snap_grid.Snap(old_lower_left);
                var new_pin_position = ArrangeHelper.GetPinPositionForCorner(input_xfrm, new_lower_left, corner);
                var output_xfrm = new VisioAutomation.Drawing.Point(new_pin_position.X, new_pin_position.Y);
                output_xfrms.Add(output_xfrm);
            }

            ModifyPinPositions(target.Page, target.ShapeIDs, output_xfrms);
        }


        private static Drawing.Point GetPinPositionForCorner(XForm input_xfrm, Drawing.Point new_lower_left, VisioAutomation.Scripting.Layout.SnapCornerPosition corner)
        {
            var size = new Drawing.Size(input_xfrm.Width, input_xfrm.Height);
            var locpin = new Drawing.Point(input_xfrm.LocPinX, input_xfrm.LocPinY);

            switch (corner)
            {
                case VisioAutomation.Scripting.Layout.SnapCornerPosition.LowerLeft:
                    {
                        return new_lower_left.Add(locpin.X, locpin.Y);
                    }
                case VisioAutomation.Scripting.Layout.SnapCornerPosition.UpperRight:
                    {
                        return new_lower_left.Subtract(size.Width, size.Height).Add(locpin.X, locpin.Y);
                    }
                case VisioAutomation.Scripting.Layout.SnapCornerPosition.LowerRight:
                    {
                        return new_lower_left.Subtract(size.Width, 0).Add(locpin.X, locpin.Y);
                    }
                case VisioAutomation.Scripting.Layout.SnapCornerPosition.UpperLeft:
                    {
                        return new_lower_left.Subtract(0, size.Height).Add(locpin.X, locpin.Y);
                    }
                default:
                    {
                        throw new System.ArgumentOutOfRangeException(nameof(corner), "Unsupported corner");
                    }
            }
        }

        public static void SnapSize(TargetShapeIDs target, Drawing.Size snapsize, Drawing.Size minsize)
        {
            var input_xfrms = XForm.Get(target);
            var sizes = new List<VisioAutomation.Drawing.Size>(input_xfrms.Count);

            var grid = new SnappingGrid(snapsize);
            foreach (var input_xfrm in input_xfrms)
            {
                // First snap the size to the grid
                double old_w = input_xfrm.Width;
                double old_h = input_xfrm.Height;
                var input_size = new Drawing.Size(old_w, old_h);
                var snapped_size = grid.Snap(input_size);

                // then account for any minum size requirements
                double new_w = System.Math.Max(snapped_size.Width, minsize.Width);
                double new_h = System.Math.Max(snapped_size.Height, minsize.Height);

                sizes.Add(new VisioAutomation.Drawing.Size(new_w, new_h));
            }

            // Now apply the updates to the sizes
            ModifySizes(target.Page, target.ShapeIDs, sizes);
        }
    }
}
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Helpers
{
    internal static class ArrangeHelper
    {
        private static double _get_position_on_shape(Models.ShapeXFormData xform, Models.ShapeRelativePosition pos)
        {
            var r = xform.GetRectangle();

            switch (pos)
            {
                case VisioScripting.Models.ShapeRelativePosition.PinY:
                    return xform.XFormPinY;
                case VisioScripting.Models.ShapeRelativePosition.PinX:
                    return xform.XFormPinX;
                case VisioScripting.Models.ShapeRelativePosition.Left:
                    return r.Left;
                case VisioScripting.Models.ShapeRelativePosition.Right:
                    return r.Right;
                case VisioScripting.Models.ShapeRelativePosition.Top:
                    return r.Top;
                case VisioScripting.Models.ShapeRelativePosition.Bottom:
                    return r.Bottom;
            }

            throw new System.ArgumentOutOfRangeException(nameof(pos));
        }

        internal static List<int> _sort_shapes_by_position(IVisio.Page page, IList<int> shapeids, Models.ShapeRelativePosition pos)
        {
            // First get the transforms of the shapes on the given axis
            var xforms = VisioScripting.Models.ShapeXFormData._get_xfrms(page, shapeids);

            // Then, sort the shapeids pased on the corresponding value in the results

            var sorted_shapeids = Enumerable.Range(0, shapeids.Count)
                .Select(i => new { index = i, shapeid = shapeids[i], pos = ArrangeHelper._get_position_on_shape(xforms[i], pos) })
                .OrderBy(i => i.pos)
                .Select(i => i.shapeid)
                .ToList();

            return sorted_shapeids;
        }

        internal static void _distribute_with_spacing(IVisio.Page page, IList<int> shapeids, Models.Axis axis, double spacing)
        {
            if (spacing < 0.0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(spacing));
            }

            if (shapeids.Count < 2)
            {
                return;
            }

            // Calculate the new Xfrms
            var sortpos = axis == VisioScripting.Models.Axis.XAxis
                ? VisioScripting.Models.ShapeRelativePosition.PinX
                : VisioScripting.Models.ShapeRelativePosition.PinY;

            var delta = axis == VisioScripting.Models.Axis.XAxis
                ? new VisioAutomation.Geometry.Size(spacing, 0)
                : new VisioAutomation.Geometry.Size(0, spacing);


            var input_xfrms = VisioScripting.Models.ShapeXFormData._get_xfrms(page, shapeids);
            var bb = VisioScripting.Models.ShapeXFormData.GetBoundingBox(input_xfrms);
            var cur_pos = new VisioAutomation.Geometry.Point(bb.Left, bb.Bottom);

            var newpositions = new List<VisioAutomation.Geometry.Point>(shapeids.Count);
            foreach (var input_xfrm in input_xfrms)
            {
                var new_pinpos = axis == VisioScripting.Models.Axis.XAxis
                    ? new VisioAutomation.Geometry.Point(cur_pos.X + input_xfrm.XFormLocPinX, input_xfrm.XFormPinY)
                    : new VisioAutomation.Geometry.Point(input_xfrm.XFormPinX, cur_pos.Y + input_xfrm.XFormLocPinY);

                newpositions.Add(new_pinpos);
                cur_pos = cur_pos.Add(input_xfrm.XFormWidth, input_xfrm.XFormHeight).Add(delta);
            }

            // Apply the changes
            var sorted_shapeids = ArrangeHelper._sort_shapes_by_position(page, shapeids, sortpos);

            _modify_pin_positions(page, sorted_shapeids, newpositions);
        }

        private static void _modify_pin_positions(IVisio.Page page, IList<int> sorted_shapeids, List<VisioAutomation.Geometry.Point> newpositions)
        {
            var writer = new SidSrcWriter();
            for (int i = 0; i < newpositions.Count; i++)
            {
                writer.SetValue((short)sorted_shapeids[i], VisioAutomation.ShapeSheet.SrcConstants.XFormPinX, newpositions[i].X);
                writer.SetValue((short)sorted_shapeids[i], VisioAutomation.ShapeSheet.SrcConstants.XFormPinY, newpositions[i].Y);
            }

            writer.Commit(page, VisioAutomation.ShapeSheet.CellValueType.Formula);
        }
    }
}
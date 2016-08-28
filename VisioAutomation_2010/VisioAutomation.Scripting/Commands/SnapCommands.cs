using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Scripting.Layout;
using VisioAutomation.Scripting.Utilities;

namespace VisioAutomation.Scripting.Commands
{
    public class SnapCommands : CommandSet
    {
        internal SnapCommands(Client client) :
            base(client)
        {
        }

        public void SnapSize(TargetShapes targets, double w, double h)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapes2DOnly(this._client);
            if (shapes.Count < 1)
            {
                return;
            }


            var shapeids = shapes.Select(s => s.ID).ToList();

            var application = this._client.Application.Get();
            var target_ids = targets.ToShapeIDs(application.ActivePage);
            using (var undoscope = this._client.Application.NewUndoScope("Snape Shape Sizes"))
            {
                var snapsize = new Drawing.Size(w, h);
                var minsize = new Drawing.Size(w, h);
                SnapCommands.SnapSize(target_ids, snapsize, minsize);
            }
        }

        private static void SnapSize(TargetShapeIDs target, Drawing.Size snapsize, Drawing.Size minsize)
        {
            var input_xfrms = Shapes.XFormCells.GetCells(target.Page, target.ShapeIDs);
            var output_xfrms = new List<Shapes.XFormCells>(input_xfrms.Count);

            var grid = new SnappingGrid(snapsize);
            foreach (var input_xfrm in input_xfrms)
            {
                var inut_size = new Drawing.Size(input_xfrm.Width.Result, input_xfrm.Height.Result);
                var snapped_size = grid.Snap(inut_size);
                double max_w = System.Math.Max(snapped_size.Width, minsize.Width);
                double max_h = System.Math.Max(snapped_size.Height, minsize.Height);
                var new_size = new Drawing.Size(max_w, max_h);

                var output_xfrm = new Shapes.XFormCells();
                output_xfrm.Width = new_size.Width;
                output_xfrm.Height = new_size.Height;

                output_xfrms.Add(output_xfrm);
            }

            // Now apply them
            ArrangeHelper.update_xfrms(target, output_xfrms);
        }


        public void SnapCorner(TargetShapes targets, double w, double h, SnapCornerPosition corner)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapes2DOnly(this._client);

            if (shapes.Count < 1)
            {
                return;
            }

            var application = this._client.Application.Get();
            var target_ids = targets.ToShapeIDs(application.ActivePage);
            using (var undoscope = this._client.Application.NewUndoScope("SnapCorner"))
            {
                ArrangeHelper.SnapCorner(target_ids, new Drawing.Size(w, h), corner);
            }
        }

        public void SnapSize(TargetShapes targets, Drawing.Size snapsize, Drawing.Size minsize)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapes2DOnly(this._client);

            if (shapes.Count < 1)
            {
                return;
            }

            var application = this._client.Application.Get();
            var target_ids = targets.ToShapeIDs(application.ActivePage);
            using (var undoscope = this._client.Application.NewUndoScope("SnapSize"))
            {
                ArrangeHelper.SnapSize(target_ids, snapsize, minsize);
            }
        }
    }
}
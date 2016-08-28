using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using VisioAutomation.Drawing.Layout;
using VisioAutomation.Scripting.Exceptions;
using VisioAutomation.Scripting.Layout;
using VisioAutomation.Scripting.Utilities;
using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomation.Scripting.Commands
{
    public class ArrangeCommands : CommandSet
    {
        internal ArrangeCommands(Client client) :
            base(client)
        {

        }





        public void Nudge(TargetShapes targets, double dx, double dy)
        {
            if (dx == 0.0 && dy == 0.0)
            {
                return;
            }

            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 1)
            {
                return;
            }

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Nudge Shapes"))
            {
                var selection = this._client.Selection.Get();
                var unitcode = IVisio.VisUnitCodes.visInches;

                // Move method: http://msdn.microsoft.com/en-us/library/ms367549.aspx   
                selection.Move(dx, dy, unitcode);
            }
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
            using (var undoscope = this._client.Application.NewUndoScope("Snape Shape Sizes"))
            {
                var active_page = application.ActivePage;
                var snapsize = new Drawing.Size(w, h);
                var minsize = new Drawing.Size(w, h);
                ArrangeCommands.SnapSize(active_page, shapeids, snapsize, minsize);
            }
        }

        private static void SnapSize(IVisio.Page page, IList<int> shapeids, Drawing.Size snapsize, Drawing.Size minsize)
        {
            var input_xfrms = Shapes.XFormCells.GetCells(page, shapeids);
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
            ArrangeHelper.update_xfrms(page, shapeids, output_xfrms);
        }

        public void Send(TargetShapes targets, Selections.ShapeSendDirection dir)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 1)
            {
                return;
            }

            var selection = this._client.Selection.Get();
            Selections.SelectionHelper.SendShapes(selection, dir);
        }


        public IList<Shapes.XFormCells> GetXForm(TargetShapes targets)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapes(this._client);
            if (shapes.Count<1)
            {
                return new List<Shapes.XFormCells>(0);
            }

            var shapeids = shapes.Select(s=>s.ID).ToList();
            var application = this._client.Application.Get();
            var page = application.ActivePage;
            var data = Shapes.XFormCells.GetCells(page, shapeids);
            return data;
        }

        public IVisio.Shape Group()
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            // No shapes provided, use the active selection
            if (!this._client.Selection.HasShapes())
            {
                throw new VisioOperationException("No Selected Shapes to Group");
            }

            // the other way of doing this: this.Client.VisioApplication.DoCmd((short)IVisio.VisUICmds.visCmdObjectGroup);
            // but it doesn't return the group

            var selection = this._client.Selection.Get();
            var g = selection.Group();
            return g;
        }

        public void Ungroup(TargetShapes targets)
        {
            this._client.Application.AssertApplicationAvailable();
            if (targets.Shapes == null)
            {
                if (this._client.Selection.HasShapes())
                {
                    var application = this._client.Application.Get();
                    application.DoCmd((short)IVisio.VisUICmds.visCmdObjectUngroup);
                }
            }
            else
            {
                foreach (var shape in targets.Shapes)
                {
                    shape.Ungroup();
                }
            }
        }

        public void SetLock(TargetShapes targets, Shapes.LockCells lockcells)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapes(this._client);
            if (shapes.Count < 1)
            {
                return;
            } 

            var selection = this._client.Selection.Get();
            var shapeids = selection.GetIDs();
            var update = new FormulaWriterSIDSRC();

            foreach (int shapeid in shapeids)
            {
                lockcells.SetFormulas((short)shapeid, update);
            }

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Set Shape Lock Properties"))
            {
                var active_page = application.ActivePage;
                update.Commit(active_page);
            }
        }

        public void SetSize(TargetShapes targets, double? w, double? h)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapes(this._client);
            if (shapes.Count < 1)
            {
                return;
            } 

            var shapeids = shapes.Select(s=>s.ID).ToList();
            var update = new FormulaWriterSIDSRC();
            foreach (int shapeid in shapeids)
            {
                if (w.HasValue && w.Value>=0)
                {
                    update.SetFormula((short)shapeid, ShapeSheet.SRCConstants.Width, w.Value);
                }
                if (h.HasValue && h.Value >= 0)
                {
                    update.SetFormula((short)shapeid, ShapeSheet.SRCConstants.Height, h.Value);                    
                }
            }

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Set Shape Size"))
            {
                var active_page = application.ActivePage;
                update.Commit(active_page);
            }
        }

        public static Drawing.Rectangle GetBoundingBox(IEnumerable<Shapes.XFormCells> xfrms)
        {
            var bb = new BoundingBox(xfrms.Select(ArrangeHelper.GetRectangle));
            if (!bb.HasValue)
            {
                throw new System.ArgumentException("Could not calculate bounding box");
            }
            return bb.Rectangle;
        }

        public void Stack(Axis axis, double space)
        {
            if (!this._client.Selection.HasShapes(2))
            {
                return;
            }

            if (space < 0.0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(space), "must be non-negative");
            }

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Stack"))
            {
                var targets = new TargetShapes();
                if (axis == Axis.YAxis)
                {
                    this._client.Align.AlignHorizontal(targets,AlignmentHorizontal.Center);
                }
                else
                {
                    this._client.Align.AlignVertical(targets,AlignmentVertical.Center);
                }
                this.DistributeOnAxis(axis, space);
            }
        }

        public void DistributeOnAxis(Axis axis, double d)
        {
            if (!this._client.Document.HasActiveDocument)
            {
                return;
            }
            var application = this._client.Application.Get();
            var selection = this._client.Selection.Get();
            var shapeids = selection.GetIDs();
            using (var undoscope = this._client.Application.NewUndoScope("Distribute on Axis"))
            {
                ArrangeHelper.DistributeWithSpacing(application.ActivePage, shapeids, axis, d);
            }
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

            var shapeids = shapes.Select(s => s.ID).ToList();
            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("SnapCorner"))
            {
                var active_page = application.ActivePage;
                ArrangeHelper.SnapCorner(active_page, shapeids, new Drawing.Size(w, h), corner);
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

            var shapeids = shapes.Select(s => s.ID).ToList();
            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("SnapSize"))
            {
                var active_page = application.ActivePage;
                ArrangeHelper.SnapSize(active_page, shapeids, snapsize, minsize);
            }
        }
    }
}
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;

namespace VisioAutomation.Scripting.Commands
{
    public class ArrangeCommands : CommandSet
    {
        internal ArrangeCommands(Client client) :
            base(client)
        {

        }

        private static IVisio.VisUICmds _map_halign_to_uicmd(Drawing.AlignmentHorizontal v)
            {
                switch (v)
                {
                    case Drawing.AlignmentHorizontal.Left: return IVisio.VisUICmds.visCmdDistributeLeft;
                    case Drawing.AlignmentHorizontal.Center: return IVisio.VisUICmds.visCmdDistributeCenter;
                    case Drawing.AlignmentHorizontal.Right: return IVisio.VisUICmds.visCmdDistributeRight;
                    default: throw new System.ArgumentOutOfRangeException();
                }
            }

        private static IVisio.VisUICmds _map_valign_to_uicmd(Drawing.AlignmentVertical v)
        {
            switch (v)
            {
                case Drawing.AlignmentVertical.Top: return IVisio.VisUICmds.visCmdDistributeTop;
                case Drawing.AlignmentVertical.Center: return IVisio.VisUICmds.visCmdDistributeMiddle;
                case Drawing.AlignmentVertical.Bottom: return IVisio.VisUICmds.visCmdDistributeBottom;
                default: throw new System.ArgumentOutOfRangeException();
            }
        }

        private static IVisio.VisUICmds _map_axis_to_uicmd(Drawing.Axis v)
        {
            switch (v)
            {
                case Drawing.Axis.XAxis: return IVisio.VisUICmds.visCmdDistributeHSpace;
                case Drawing.Axis.YAxis: return IVisio.VisUICmds.visCmdDistributeVSpace;
                default: throw new System.ArgumentOutOfRangeException();
            }
        }

        private static IVisio.VisVerticalAlignTypes _map_isd_valign_to_visio_valign(Drawing.AlignmentVertical v)
        {
            switch (v)
            {
                case Drawing.AlignmentVertical.Top: return IVisio.VisVerticalAlignTypes.visVertAlignTop;
                case Drawing.AlignmentVertical.Center: return IVisio.VisVerticalAlignTypes.visVertAlignMiddle;
                case Drawing.AlignmentVertical.Bottom: return IVisio.VisVerticalAlignTypes.visVertAlignBottom;
                default: throw new System.ArgumentOutOfRangeException();
            }
        }

        private static IVisio.VisHorizontalAlignTypes _map_isd_halign_to_visio_halign(Drawing.AlignmentHorizontal v)
        {
            switch (v)
            {
                case Drawing.AlignmentHorizontal.Left: return IVisio.VisHorizontalAlignTypes.visHorzAlignLeft;
                case Drawing.AlignmentHorizontal.Center: return IVisio.VisHorizontalAlignTypes.visHorzAlignCenter;
                case Drawing.AlignmentHorizontal.Right: return IVisio.VisHorizontalAlignTypes.visHorzAlignRight;
                default: throw new System.ArgumentOutOfRangeException();
            }
        }

        public void Distribute(IList<IVisio.Shape> target_shapes, Drawing.AlignmentHorizontal halign)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            int shape_count = this.GetTargetSelection(target_shapes);
            if (shape_count < 1)
            {
                return;
            }

            var cmd = ArrangeCommands._map_halign_to_uicmd(halign);

            var application = this.Client.Application.Get();
            application.DoCmd((short)cmd);
        }

        public void Distribute(IList<IVisio.Shape> target_shapes, Drawing.AlignmentVertical valign)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            int shape_count = this.GetTargetSelection(target_shapes);
            if (shape_count < 1)
            {
                return;
            }

            var cmd = ArrangeCommands._map_valign_to_uicmd(valign);

            var application = this.Client.Application.Get();
            application.DoCmd((short)cmd); 
        }

        public void Distribute(IList<IVisio.Shape> target_shapes, Drawing.Axis axis)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            int shape_count = this.GetTargetSelection(target_shapes);
            if (shape_count < 1)
            {
                return;
            }

            var cmd = ArrangeCommands._map_axis_to_uicmd(axis);

            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Distribute Shapes"))
            {
                application.DoCmd((short)cmd);
            }
        }

        public void Nudge(IList<IVisio.Shape> target_shapes, double dx, double dy)
        {
            if (dx == 0.0 && dy == 0.0)
            {
                return;
            }

            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            int shape_count = this.GetTargetSelection(target_shapes);
            if (shape_count < 1)
            {
                return;
            }

            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Nudge Shapes"))
            {
                var selection = this.Client.Selection.Get();
                var unitcode = IVisio.VisUnitCodes.visInches;

                // Move method: http://msdn.microsoft.com/en-us/library/ms367549.aspx   
                selection.Move(dx, dy, unitcode);
            }
        }

        public void SnapSize(IList<IVisio.Shape> target_shapes, double w, double h)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();
            
            var shapes = this.GetTargetShapes2D(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }


            var shapeids = shapes.Select(s => s.ID).ToList();

            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Snape Shape Sizes"))
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

            var grid = new Drawing.SnappingGrid(snapsize);
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
            ArrangeCommands.update_xfrms(page, shapeids, output_xfrms);
        }

        private static void update_xfrms(IVisio.Page page, IList<int> shapeids, IList<Shapes.XFormCells> xfrms)
        {

            var update = new ShapeSheet.Update();
            for (int i = 0; i < shapeids.Count; i++)
            {
                var shape_id = shapeids[i];
                var xfrm = xfrms[i];
                update.SetFormulas((short)shape_id, xfrm);
            }
            update.Execute(page);
        }



        public void Send(IList<IVisio.Shape> target_shapes, Selection.ShapeSendDirection dir)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            int shape_count = this.GetTargetSelection(target_shapes);
            if (shape_count < 1)
            {
                return;
            }

            var selection = this.Client.Selection.Get();
            Selection.SelectionHelper.SendShapes(selection, dir);
        }

        public void Align(IList<IVisio.Shape> target_shapes, Drawing.AlignmentHorizontal align)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            int shape_count = this.GetTargetSelection(target_shapes);
            if (shape_count < 2)
            {
                return;
            }

            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Align Shapes"))
            {
                const bool glue_to_guide = false;
                var selection = this.Client.Selection.Get();
                var halign = ArrangeCommands._map_isd_halign_to_visio_halign(align);
                var valign = IVisio.VisVerticalAlignTypes.visVertAlignNone;
                selection.Align(halign, valign, glue_to_guide);
            }
        }

        public void Align(IList<IVisio.Shape> target_shapes, Drawing.AlignmentVertical align)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            int shape_count = this.GetTargetSelection(target_shapes);
            if (shape_count < 2)
            {
                return;
            }

            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Align Shapes"))
            {
                bool glue_to_guide = false;
                var selection = this.Client.Selection.Get();
                var halign = IVisio.VisHorizontalAlignTypes.visHorzAlignNone;

                var valign = ArrangeCommands._map_isd_valign_to_visio_valign(align);

                selection.Align(halign, valign, glue_to_guide);
            }
        }

        public IList<Shapes.XFormCells> GetXForm(IList<IVisio.Shape> target_shapes)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count<1)
            {
                return new List<Shapes.XFormCells>(0);
            }

            var shapeids = shapes.Select(s=>s.ID).ToList();
            var application = this.Client.Application.Get();
            var page = application.ActivePage;
            var data = Shapes.XFormCells.GetCells(page, shapeids);
            return data;
        }

        public IVisio.Shape Group()
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            // No shapes provided, use the active selection
            if (!this.Client.Selection.HasShapes())
            {
                throw new VisioOperationException("No Selected Shapes to Group");
            }

            // the other way of doing this: this.Client.VisioApplication.DoCmd((short)IVisio.VisUICmds.visCmdObjectGroup);
            // but it doesn't return the group

            var selection = this.Client.Selection.Get();
            var g = selection.Group();
            return g;
        }

        public void Ungroup(IList<IVisio.Shape> target_shapes)
        {
            this.Client.Application.AssertApplicationAvailable();
            if (target_shapes == null)
            {
                if (this.Client.Selection.HasShapes())
                {
                    var application = this.Client.Application.Get();
                    application.DoCmd((short)IVisio.VisUICmds.visCmdObjectUngroup);
                }
            }
            else
            {
                foreach (var shape in target_shapes)
                {
                    shape.Ungroup();
                }
            }
        }

        public void SetLock(IList<IVisio.Shape> target_shapes, Shapes.LockCells lockcells)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            } 

            var selection = this.Client.Selection.Get();
            var shapeids = selection.GetIDs();
            var update = new ShapeSheet.Update();

            foreach (int shapeid in shapeids)
            {
                update.SetFormulas((short)shapeid, lockcells);
            }

            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Set Shape Lock Properties"))
            {
                var active_page = application.ActivePage;
                update.Execute(active_page);
            }
        }

        public void SetSize(IList<IVisio.Shape> target_shapes, double? w, double? h)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            } 

            var shapeids = shapes.Select(s=>s.ID).ToList();
            var update = new ShapeSheet.Update();
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

            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Set Shape Size"))
            {
                var active_page = application.ActivePage;
                update.Execute(active_page);
            }
        }

        private static Drawing.Rectangle GetRectangle(Shapes.XFormCells xFormCells)
        {
            var pin = new Drawing.Point(xFormCells.PinX.Result, xFormCells.PinY.Result);
            var locpin = new Drawing.Point(xFormCells.LocPinX.Result, xFormCells.LocPinY.Result);
            var size = new Drawing.Size(xFormCells.Width.Result, xFormCells.Height.Result);
            return new Drawing.Rectangle(pin - locpin, size);
        }

        public static Drawing.Rectangle GetBoundingBox(IEnumerable<Shapes.XFormCells> xfrms)
        {
            var bb = new Drawing.BoundingBox(xfrms.Select(ArrangeCommands.GetRectangle));
            if (!bb.HasValue)
            {
                throw new System.ArgumentException("Could not calculate bounding box");
            }
            else
            {
                return bb.Rectangle;
            }
        }

        public void Stack(Drawing.Axis axis, double space)
        {


            if (!this.Client.Selection.HasShapes(2))
            {
                return;
            }
            if (space < 0.0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(space), "must be non-negative");
            }

            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Stack"))
            {
                if (axis == Drawing.Axis.YAxis)
                {
                    this.Align(null,Drawing.AlignmentHorizontal.Center);
                }
                else
                {
                    this.Align(null,Drawing.AlignmentVertical.Center);
                }
                this.Distribute(axis, space);
            }
        }

        public void Distribute(Drawing.Axis axis, double d)
        {
            if (!this.Client.Document.HasActiveDocument)
            {
                return;
            }
            var application = this.Client.Application.Get();
            var selection = this.Client.Selection.Get();
            var shapeids = selection.GetIDs();
            using (var undoscope = this.Client.Application.NewUndoScope("Distribute"))
            {
                ArrangeHelper.DistributeWithSpacing(application.ActivePage, shapeids, axis, d);
            }
        }
        
        public void SnapCorner(IList<IVisio.Shape> target_shapes, double w, double h, SnapCornerPosition corner)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes2D(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }

            var shapeids = shapes.Select(s => s.ID).ToList();
            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("SnapCorner"))
            {
                var active_page = application.ActivePage;
                ArrangeHelper.SnapCorner(active_page, shapeids, new Drawing.Size(w, h), corner);
            }
        }

        public void SnapSize(IList<IVisio.Shape> target_shapes, Drawing.Size snapsize, Drawing.Size minsize)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes2D(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }

            var shapeids = shapes.Select(s => s.ID).ToList();
            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("SnapSize"))
            {
                var active_page = application.ActivePage;
                ArrangeHelper.SnapSize(active_page, shapeids, snapsize, minsize);
            }
        }
    }
}
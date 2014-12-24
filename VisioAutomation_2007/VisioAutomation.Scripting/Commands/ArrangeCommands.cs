using System.Linq;
using VisioAutomation.Application;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.Scripting.Commands
{
    public class ArrangeCommands : CommandSet
    {
        public ArrangeCommands(Client client) :
            base(client)
        {

        }

        private static IVisio.VisUICmds _map_halign_to_uicmd(VA.Drawing.AlignmentHorizontal v)
            {
                switch (v)
                {
                    case VA.Drawing.AlignmentHorizontal.Left: return IVisio.VisUICmds.visCmdDistributeLeft;
                    case VA.Drawing.AlignmentHorizontal.Center: return IVisio.VisUICmds.visCmdDistributeCenter;
                    case VA.Drawing.AlignmentHorizontal.Right: return IVisio.VisUICmds.visCmdDistributeRight;
                    default: throw new System.ArgumentOutOfRangeException();
                }
            }

        private static IVisio.VisUICmds _map_valign_to_uicmd(VA.Drawing.AlignmentVertical v)
        {
            switch (v)
            {
                case VA.Drawing.AlignmentVertical.Top: return IVisio.VisUICmds.visCmdDistributeTop;
                case VA.Drawing.AlignmentVertical.Center: return IVisio.VisUICmds.visCmdDistributeMiddle;
                case VA.Drawing.AlignmentVertical.Bottom: return IVisio.VisUICmds.visCmdDistributeBottom;
                default: throw new System.ArgumentOutOfRangeException();
            }
        }

        private static IVisio.VisUICmds _map_axis_to_uicmd(VA.Drawing.Axis v)
        {
            switch (v)
            {
                case VA.Drawing.Axis.XAxis: return IVisio.VisUICmds.visCmdDistributeHSpace;
                case VA.Drawing.Axis.YAxis: return IVisio.VisUICmds.visCmdDistributeVSpace;
                default: throw new System.ArgumentOutOfRangeException();
            }
        }

        private static IVisio.VisVerticalAlignTypes _map_isd_valign_to_visio_valign(VA.Drawing.AlignmentVertical v)
        {
            switch (v)
            {
                case VA.Drawing.AlignmentVertical.Top: return IVisio.VisVerticalAlignTypes.visVertAlignTop;
                case VA.Drawing.AlignmentVertical.Center: return IVisio.VisVerticalAlignTypes.visVertAlignMiddle;
                case VA.Drawing.AlignmentVertical.Bottom: return IVisio.VisVerticalAlignTypes.visVertAlignBottom;
                default: throw new System.ArgumentOutOfRangeException();
            }
        }

        private static IVisio.VisHorizontalAlignTypes _map_isd_halign_to_visio_halign(VA.Drawing.AlignmentHorizontal v)
        {
            switch (v)
            {
                case VA.Drawing.AlignmentHorizontal.Left: return IVisio.VisHorizontalAlignTypes.visHorzAlignLeft;
                case VA.Drawing.AlignmentHorizontal.Center: return IVisio.VisHorizontalAlignTypes.visHorzAlignCenter;
                case VA.Drawing.AlignmentHorizontal.Right: return IVisio.VisHorizontalAlignTypes.visHorzAlignRight;
                default: throw new System.ArgumentOutOfRangeException();
            }
        }

        public void Distribute(IList<IVisio.Shape> target_shapes, VA.Drawing.AlignmentHorizontal halign)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            int shape_count = this.GetTargetSelection(target_shapes);
            if (shape_count < 1)
            {
                return;
            }

            var cmd = _map_halign_to_uicmd(halign);

            this.Client.VisioApplication.DoCmd((short)cmd);
        }

        public void Distribute(IList<IVisio.Shape> target_shapes, VA.Drawing.AlignmentVertical valign)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            int shape_count = this.GetTargetSelection(target_shapes);
            if (shape_count < 1)
            {
                return;
            }

            var cmd = _map_valign_to_uicmd(valign);

            this.Client.VisioApplication.DoCmd((short)cmd); 
        }

        public void Distribute(IList<IVisio.Shape> target_shapes, VA.Drawing.Axis axis)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            int shape_count = this.GetTargetSelection(target_shapes);
            if (shape_count < 1)
            {
                return;
            }

            var cmd = _map_axis_to_uicmd(axis);

            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication,"Distribute Shapes"))
            {
                this.Client.VisioApplication.DoCmd((short)cmd);
            }
        }

        public void Nudge(IList<IVisio.Shape> target_shapes, double dx, double dy)
        {
            if (dx == 0.0 && dy == 0.0)
            {
                return;
            }

            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            int shape_count = this.GetTargetSelection(target_shapes);
            if (shape_count < 1)
            {
                return;
            }

            var application = this.Client.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication,"Nudge Shapes"))
            {
                var selection = this.Client.Selection.Get();
                var unitcode = IVisio.VisUnitCodes.visInches;

                // Move method: http://msdn.microsoft.com/en-us/library/ms367549.aspx   
                selection.Move(dx, dy, unitcode);
            }
        }


                
        public void SnapSize(IList<IVisio.Shape> target_shapes, double w, double h)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();
            
            var shapes = GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }

            var shapes_2d = shapes.Where(s=>s.OneD==0).ToList();
            var shapeids = shapes_2d.Select(s => s.ID).ToList();

            var application = this.Client.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication,"Snape Shape Sizes"))
            {
                var active_page = application.ActivePage;
                var snapsize = new VA.Drawing.Size(w, h);
                var minsize = new VA.Drawing.Size(w, h);
                SnapSize(active_page, shapeids, snapsize, minsize);
            }
        }


        private static void SnapSize(IVisio.Page page, IList<int> shapeids, VA.Drawing.Size snapsize, VA.Drawing.Size minsize)
        {
            var input_xfrms = VA.Shapes.XFormCells.GetCells(page, shapeids);
            var output_xfrms = new List<VA.Shapes.XFormCells>(input_xfrms.Count);

            var grid = new VA.Drawing.SnappingGrid(snapsize);
            foreach (var input_xfrm in input_xfrms)
            {
                var inut_size = new VA.Drawing.Size(input_xfrm.Width.Result, input_xfrm.Height.Result);
                var snapped_size = grid.Snap(inut_size);
                double max_w = System.Math.Max(snapped_size.Width, minsize.Width);
                double max_h = System.Math.Max(snapped_size.Height, minsize.Height);
                var new_size = new VA.Drawing.Size(max_w, max_h);

                var output_xfrm = new VA.Shapes.XFormCells();
                output_xfrm.Width = new_size.Width;
                output_xfrm.Height = new_size.Height;

                output_xfrms.Add(output_xfrm);
            }

            // Now apply them
            update_xfrms(page, shapeids, output_xfrms);
        }

        private static void update_xfrms(IVisio.Page page, IList<int> shapeids, IList<VA.Shapes.XFormCells> xfrms)
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



        public void Send(IList<IVisio.Shape> target_shapes, VA.Selection.ShapeSendDirection dir)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            int shape_count = this.GetTargetSelection(target_shapes);
            if (shape_count < 1)
            {
                return;
            }

            var selection = Client.Selection.Get();
            VA.Selection.SelectionHelper.SendShapes(selection, dir);
        }

        public void Align(IList<IVisio.Shape> target_shapes, VA.Drawing.AlignmentHorizontal align)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            int shape_count = this.GetTargetSelection(target_shapes);
            if (shape_count < 2)
            {
                return;
            }

            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication,"Align Shapes"))
            {
                bool glue_to_guide = false;
                var selection = Client.Selection.Get();
                var halign = _map_isd_halign_to_visio_halign(align);
                var valign = IVisio.VisVerticalAlignTypes.visVertAlignNone;
                selection.Align(halign, valign, glue_to_guide);
            }
        }

        public void Align(IList<IVisio.Shape> target_shapes, VA.Drawing.AlignmentVertical align)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            int shape_count = this.GetTargetSelection(target_shapes);
            if (shape_count < 2)
            {
                return;
            }

            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication,"Align Shapes"))
            {
                bool glue_to_guide = false;
                var selection = Client.Selection.Get();
                var halign = IVisio.VisHorizontalAlignTypes.visHorzAlignNone;

                var valign = _map_isd_valign_to_visio_valign(align);

                selection.Align(halign, valign, glue_to_guide);
            }
        }

        public IList<VA.Shapes.XFormCells> GetXForm(IList<IVisio.Shape> target_shapes)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count<1)
            {
                return new List<VA.Shapes.XFormCells>(0);
            }

            var shapeids = shapes.Select(s=>s.ID).ToList();
            var page = this.Client.VisioApplication.ActivePage;
            var data = VA.Shapes.XFormCells.GetCells(page, shapeids);
            return data;
        }

        public IVisio.Shape Group()
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            // No shapes provided, use the active selection
            if (!this.Client.HasSelectedShapes())
            {
                throw new ScriptingException("No Selected Shapes to Group");
            }

            // the other way of doing this: this.Client.VisioApplication.DoCmd((short)IVisio.VisUICmds.visCmdObjectGroup);
            // but it doesn't return the group

            var selection = this.Client.Selection.Get();
            var g = selection.Group();
            return g;
        }

        public void Ungroup(IList<IVisio.Shape> target_shapes)
        {
            this.AssertApplicationAvailable();
            if (target_shapes == null)
            {
                if (this.Client.HasSelectedShapes())
                {
                    this.Client.VisioApplication.DoCmd((short)IVisio.VisUICmds.visCmdObjectUngroup);
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

        public void SetLock(IList<IVisio.Shape> target_shapes, VA.Shapes.LockCells lockcells)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            var shapes = GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            } 

            var selection = this.Client.Selection.Get();
            var shapeids = selection.GetIDs();
            var update = new VA.ShapeSheet.Update();

            foreach (int shapeid in shapeids)
            {
                update.SetFormulas((short)shapeid, lockcells);
            }

            var application = this.Client.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication,"Set Shape Lock Properties"))
            {
                var active_page = application.ActivePage;
                update.Execute(active_page);
            }
        }

        public void SetSize(IList<IVisio.Shape> target_shapes, double? w, double? h)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            var shapes = GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            } 

            var shapeids = shapes.Select(s=>s.ID).ToList();
            var update = new VA.ShapeSheet.Update();
            foreach (int shapeid in shapeids)
            {
                if (w.HasValue && w.Value>=0)
                {
                    update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.Width, w.Value);
                }
                if (h.HasValue && h.Value >= 0)
                {
                    update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.Height, h.Value);                    
                }
            }

            var application = this.Client.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication,"Set Shape Size"))
            {
                var active_page = application.ActivePage;
                update.Execute(active_page);
            }
        }

        private static VA.Drawing.Rectangle GetRectangle(VA.Shapes.XFormCells xFormCells)
        {
            var pin = new VA.Drawing.Point(xFormCells.PinX.Result, xFormCells.PinY.Result);
            var locpin = new VA.Drawing.Point(xFormCells.LocPinX.Result, xFormCells.LocPinY.Result);
            var size = new VA.Drawing.Size(xFormCells.Width.Result, xFormCells.Height.Result);
            return new VA.Drawing.Rectangle(pin - locpin, size);
        }

        public static VA.Drawing.Rectangle GetBoundingBox(IEnumerable<VA.Shapes.XFormCells> xfrms)
        {
            var bb = new VA.Drawing.BoundingBox(xfrms.Select(GetRectangle));
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


            if (!this.Client.HasSelectedShapes(2))
            {
                return;
            }
            if (space < 0.0)
            {
                throw new System.ArgumentOutOfRangeException("space", "must be non-negative");
            }

            var application = this.Client.VisioApplication;
            using (var undoscope = new UndoScope(application,"Stack"))
            {
                if (axis == VA.Drawing.Axis.YAxis)
                {
                    Align(null,VA.Drawing.AlignmentHorizontal.Center);
                }
                else
                {
                    Align(null,VA.Drawing.AlignmentVertical.Center);
                }
                Distribute(axis, space);
            }
        }

        public void Distribute(VA.Drawing.Axis axis, double d)
        {
            if (!this.Client.HasActiveDocument)
            {
                return;
            }
            var application = this.Client.VisioApplication;
            var selection = this.Client.Selection.Get();
            var shapeids = selection.GetIDs();
            using (var undoscope = new UndoScope(application,"Distribute"))
            {
                ArrangeHelper.DistributeWithSpacing(application.ActivePage, shapeids, axis, d);
            }
        }



        public void SnapCorner(double w, double h, VA.Arrange.SnapCornerPosition corner)
        {
            if (!this.Client.HasSelectedShapes())
            {
                return;
            }
            var shapes_2d = Client.Selection.EnumShapes2D().ToList();
            var shapeids = shapes_2d.Select(s => s.ID).ToList();
            var application = this.Client.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(application,"SnapCorner"))
            {
                var active_page = application.ActivePage;
                ArrangeHelper.SnapCorner(active_page, shapeids, new VA.Drawing.Size(w, h), corner);
            }
        }
    }
}


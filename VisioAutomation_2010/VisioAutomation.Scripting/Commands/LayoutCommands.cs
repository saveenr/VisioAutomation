using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.Scripting.Commands
{
    public class LayoutCommands : CommandSet
    {
        public LayoutCommands(Session session) :
            base(session)
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

        public void Distribute(VA.Drawing.AlignmentHorizontal halign)
        {
            if (!this.Session.HasActiveDrawing)
            {
                return;
            }

            var cmd = _map_halign_to_uicmd(halign);

            this.Session.VisioApplication.DoCmd((short)cmd);
        }

        public void Distribute(VA.Drawing.AlignmentVertical valign)
        {
            if (!this.Session.HasActiveDrawing)
            {
                return;
            }

            var cmd = _map_valign_to_uicmd(valign);

            this.Session.VisioApplication.DoCmd((short)cmd); 
        }

        public void Distribute(VA.Drawing.Axis axis)
        {
            if (!this.Session.HasActiveDrawing)
            {
                return;
            }

            var cmd = _map_axis_to_uicmd(axis);

            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication,"Distribute Shapes"))
            {
                this.Session.VisioApplication.DoCmd((short)cmd);
            }
        }

        public void Distribute(IList<IVisio.Shape> target_shapes, VA.Drawing.Axis axis, double d)
        {
            var shapes = GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            } 

            var shapeids = shapes.Select(s => s.ID).ToList();
            var application = this.Session.VisioApplication;

            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication,"Distribute Shapes"))
            {
                VA.Layout.LayoutHelper.DistributeWithSpacing(application.ActivePage, shapeids, axis, d);
            }
        }

        public void Nudge(double dx, double dy)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication,"Nudge Shapes"))
            {
                var selection = this.Session.Selection.Get();
                var unitcode = IVisio.VisUnitCodes.visInches;

                // Move method: http://msdn.microsoft.com/en-us/library/ms367549.aspx   
                selection.Move(dx, dy, unitcode);
            }
        }

        public void SnapCorner(double w, double h, VA.Layout.SnapCornerPosition corner)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }
            var shapes_2d = Session.Selection.EnumShapes2D().ToList();
            var shapeids = shapes_2d.Select(s => s.ID).ToList();

            var application = this.Session.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication,"Snap Shape Corners"))
            {
                var active_page = application.ActivePage;
                VA.Layout.LayoutHelper.SnapCorner(active_page, shapeids, new VA.Drawing.Size(w, h), corner);
            }
        }

        public void SnapSize(IList<IVisio.Shape> target_shapes, double w, double h)
        {
            var shapes = GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }

            var shapes_2d = shapes.Where(s=>s.OneD==0).ToList();
            var shapeids = shapes_2d.Select(s => s.ID).ToList();

            var application = this.Session.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication,"Snape Shape Sizes"))
            {
                var active_page = application.ActivePage;
                var snapsize = new VA.Drawing.Size(w, h);
                var minsize = new VA.Drawing.Size(w, h);
                VA.Layout.LayoutHelper.SnapSize(active_page, shapeids, snapsize, minsize);
            }
        }

        public void Stack(VA.Drawing.Axis axis, double space)
        {
            if (!this.Session.HasSelectedShapes(2))
            {
                return;
            }
            if (space < 0.0)
            {
                throw new System.ArgumentOutOfRangeException("space", "must be non-negative");
            }

            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication,"Stack Shapes"))
            {
                if (axis == VA.Drawing.Axis.YAxis)
                {
                    Align(VA.Drawing.AlignmentHorizontal.Center);
                }
                else
                {
                    Align(VA.Drawing.AlignmentVertical.Center);
                }
                Distribute(null, axis, space);
            }
        }

        public void Send(VA.Selection.ShapeSendDirection dir)
        {

            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var selection = Session.Selection.Get();
            VA.Selection.SelectionHelper.SendShapes(selection, dir);
        }

        public void Align(VA.Drawing.AlignmentHorizontal align)
        {
            if (!this.Session.HasSelectedShapes(2))
            {
                return;
            }

            var cmd = LayoutCommands.AlignmentToUICmd(align);

            var application = this.Session.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication,"Align Shapes"))
            {
                bool glue_to_guide = false;
                var selection = Session.Selection.Get();
                var halign = _map_isd_halign_to_visio_halign(align);
                var valign = IVisio.VisVerticalAlignTypes.visVertAlignNone;
                selection.Align(halign, valign, glue_to_guide);
            }
        }

        public void Align(VA.Drawing.AlignmentVertical align)
        {
            if (!this.Session.HasSelectedShapes(2))
            {
                return;
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication,"Align Shapes"))
            {
                bool glue_to_guide = false;
                var selection = Session.Selection.Get();
                var halign = IVisio.VisHorizontalAlignTypes.visHorzAlignNone;

                var valign = _map_isd_valign_to_visio_valign(align);

                selection.Align(halign, valign, glue_to_guide);
            }
        }

        public int GetSelectedShapeCount()
        {
            if (!this.Session.HasActiveDrawing)
            {
                return 0;
            }

            var selection = Session.Selection.Get();
            return selection.Count;
        }

        public IList<VA.Layout.XFormCells> GetXForm(IList<IVisio.Shape> target_shapes)
        {
            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count<1)
            {
                return new List<VA.Layout.XFormCells>(0);
            }

            var shapeids = shapes.Select(s=>s.ID).ToList();
            var page = this.Session.VisioApplication.ActivePage;
            var data = VA.Layout.XFormCells.GetCells(page, shapeids);
            return data;
        }

        public IVisio.Shape Group(IList<IVisio.Shape> target_shapes)
        {
            if (target_shapes == null)
            {
                if (!this.Session.HasSelectedShapes())
                {
                    this.Session.VisioApplication.DoCmd((short)IVisio.VisUICmds.visCmdObjectUngroup);
                }
                else
                {
                    // do nothing
                }
            }
            else
            {
                this.Session.Selection.SelectNone();
                this.Session.Selection.Select(target_shapes);
                this.Session.VisioApplication.DoCmd((short)IVisio.VisUICmds.visCmdObjectUngroup);
            }

            var selection = this.Session.Selection.Get();
            var g = selection.Group();
            return g;
        }

        public void Ungroup(IList<IVisio.Shape> target_shapes)
        {
            if (target_shapes == null)
            {
                if (this.Session.HasSelectedShapes())
                {
                    this.Session.VisioApplication.DoCmd((short)IVisio.VisUICmds.visCmdObjectUngroup);
                }
                else
                {
                    // do nothing                    
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

        public void SetLock(IList<IVisio.Shape> target_shapes, VA.Layout.LockCells lockcells)
        {
            var shapes = GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            } 

            var selection = this.Session.Selection.Get();
            var shapeids = selection.GetIDs();
            var update = new VA.ShapeSheet.Update();

            foreach (int shapeid in shapeids)
            {
                update.SetFormulas((short)shapeid, lockcells);
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication,"Set Shape Lock Properties"))
            {
                var active_page = application.ActivePage;
                update.Execute(active_page);
            }
        }

        public void SetSize(IList<IVisio.Shape> target_shapes, double? w, double? h)
        {
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

            var application = this.Session.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication,"Set Shape Size"))
            {
                var active_page = application.ActivePage;
                update.Execute(active_page);
            }
        }

        private static IVisio.VisUICmds AlignmentToUICmd(VA.Drawing.AlignmentHorizontal a)
        {
            if (a == VA.Drawing.AlignmentHorizontal.Left)
            {
                return IVisio.VisUICmds.visCmdAlignObjectLeft;
            }
            if (a == VA.Drawing.AlignmentHorizontal.Center)
            {
                return IVisio.VisUICmds.visCmdAlignObjectCenter;
            }
            if (a == VA.Drawing.AlignmentHorizontal.Right)
            {
                return IVisio.VisUICmds.visCmdAlignObjectRight;
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }
        }

        private static IVisio.VisUICmds AlignmentToUICmd(VA.Drawing.AlignmentVertical a)
        {
            if (a == VA.Drawing.AlignmentVertical.Top) { return IVisio.VisUICmds.visCmdAlignObjectTop; }
            if (a == VA.Drawing.AlignmentVertical.Center) { return IVisio.VisUICmds.visCmdAlignObjectMiddle; }
            if (a == VA.Drawing.AlignmentVertical.Bottom) { return IVisio.VisUICmds.visCmdAlignObjectBottom; }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }
        }
    }
}
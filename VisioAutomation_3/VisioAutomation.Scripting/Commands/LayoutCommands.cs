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

            using (var undoscope = this.Session.VisioApplication.CreateUndoScope())
            {
                this.Session.VisioApplication.DoCmd((short)cmd);
            }
        }

        public void Distribute(VA.Drawing.Axis axis, double d)
        {
            if (!this.Session.HasActiveDrawing)
            {
                return;
            }


            var application = this.Session.VisioApplication;
            var selection = this.Session.Selection.Get();
            var shapeids = selection.GetIDs();

            using (var undoscope = application.CreateUndoScope())
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
            using (var undoscope = application.CreateUndoScope())
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
            using (var undoscope = application.CreateUndoScope())
            {
                var active_page = application.ActivePage;
                VA.Layout.LayoutHelper.SnapCorner(active_page, shapeids, new VA.Drawing.Size(w, h), corner);
            }
        }

        public void SnapSize(double w, double h)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var shapes_2d = this.Session.Selection.EnumShapes2D().ToList();
            var shapeids = shapes_2d.Select(s => s.ID).ToList();

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
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
            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                if (axis == VA.Drawing.Axis.YAxis)
                {
                    Align(VA.Drawing.AlignmentHorizontal.Center);
                }
                else
                {
                    Align(VA.Drawing.AlignmentVertical.Center);
                }
                Distribute(axis, space);
            }
        }

        public void Send( VA.Layout.ShapeSendDirection dir)
        {

            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var selection = Session.Selection.Get();
            VA.Layout.LayoutHelper.SendShapes(selection, dir);
        }

        public void Align(VA.Drawing.AlignmentHorizontal align)
        {
            if (!this.Session.HasSelectedShapes(2))
            {
                return;
            }

            var cmd = LayoutCommands.AlignmentToUICmd(align);

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                bool glue_to_guide = false;
                var selection = Session.Selection.Get();
                var halign = _map_isd_halign_to_visio_halign(align);
                var valign = IVisio.VisVerticalAlignTypes.visVertAlignNone;
                selection.Align(halign, valign, glue_to_guide);
            }
        }

        public void Align(VA.Drawing.AlignmentHorizontal align, double x)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var selection = Session.Selection.Get();
            var shapeids = selection.GetIDs();

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                VA.Layout.LayoutHelper.AlignTo(application.ActivePage, shapeids, align, x);
            }
        }

        public void Align(VA.Drawing.AlignmentVertical align)
        {
            if (!this.Session.HasSelectedShapes(2))
            {
                return;
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                bool glue_to_guide = false;
                var selection = Session.Selection.Get();
                var halign = IVisio.VisHorizontalAlignTypes.visHorzAlignNone;

                var valign = _map_isd_valign_to_visio_valign(align);

                selection.Align(halign, valign, glue_to_guide);
            }
        }

        public void Align(VA.Drawing.AlignmentVertical align, double y)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var selection = Session.Selection.Get();
            var shapeids = selection.GetIDs();

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                VA.Layout.LayoutHelper.AlignTo(application.ActivePage, shapeids,align,y);
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

        public IList<VA.Layout.XFormCells> GetXForm()
        {
            if (!this.Session.HasSelectedShapes())
            {
                return new List<VA.Layout.XFormCells>(0);
            }

            var selection = Session.Selection.Get();
            var shapeids = selection.GetIDs();

            var page = this.Session.VisioApplication.ActivePage;
            var data = VA.Layout.LayoutHelper.GetXForm(page, shapeids);
            return data;
        }

        public IVisio.Shape Group()
        {
            if (!this.Session.HasSelectedShapes(2))
            {
                return null;
            }

            var selection = this.Session.Selection.Get();
            var g = selection.Group();
            return g;
        }

        public void Ungroup()
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            this.Session.VisioApplication.DoCmd((short)IVisio.VisUICmds.visCmdObjectUngroup);
        }

        private void updatelock(VA.Layout.LockCells lockcells)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var selection = this.Session.Selection.Get();
            var shapeids = selection.GetIDs();
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            foreach (int shapeid in shapeids)
            {
                lockcells.Apply(update, (short) shapeid);
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                var active_page = application.ActivePage;
                update.Execute(active_page);
            }
        }

        public void LockAll()
        {
            var lockcells = new VA.Layout.LockCells();
            lockcells.SetAll("1");
            this.updatelock(lockcells);
        }

        public void UnlockAll()
        {
            var lockcells = new VA.Layout.LockCells();
            lockcells.SetAll("0");
            this.updatelock(lockcells);
        }

        public void SetWidth(double w)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var selection = this.Session.Selection.Get();
            var shapeids = selection.GetIDs();
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            foreach (int shapeid in shapeids)
            {
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.Width, w);
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                var active_page = application.ActivePage;
                update.Execute(active_page);
            }
        }

        public void SetHeight(double h)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var selection = this.Session.Selection.Get();
            var shapeids = selection.GetIDs();
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            foreach (int shapeid in shapeids)
            {
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.Height, h);
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                var active_page = application.ActivePage;
                update.Execute(active_page);
            }
        }

        public void SetSize(double w, double h)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var selection = this.Session.Selection.Get();
            var shapeids = selection.GetIDs();
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            foreach (int shapeid in shapeids)
            {
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.Width, w);
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.Height, h);
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
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
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using System.Collections.Generic;


namespace VisioAutomation.Scripting.Commands
{
    public class LayoutCommands : SessionCommands
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
            if (!this.Session.HasActiveDrawing())
            {
                return;
            }

            var cmd = _map_halign_to_uicmd(halign);

            this.Session.Application.DoCmd((short)cmd);
        }

        public void Distribute(VA.Drawing.AlignmentVertical valign)
        {
            if (!this.Session.HasActiveDrawing())
            {
                return;
            }

            var cmd = _map_valign_to_uicmd(valign);

            this.Session.Application.DoCmd((short)cmd); 
        }

        public void Distribute(VA.Drawing.Axis axis)
        {
            if (!this.Session.HasActiveDrawing())
            {
                return;
            }

            var cmd = _map_axis_to_uicmd(axis);

            using (var undoscope = this.Session.Application.CreateUndoScope())
            {
                this.Session.Application.DoCmd((short)cmd);
            }
        }

        public void Distribute(VA.Drawing.Axis axis, double d)
        {
            if (!this.Session.HasActiveDrawing())
            {
                return;
            }


            var application = this.Session.Application;
            var selection = this.Session.Selection.GetSelection();
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

            var application = this.Session.Application;
            using (var undoscope = application.CreateUndoScope())
            {
                var selection = this.Session.Selection.GetSelection();
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
            var shapes_2d = Session.Selection.EnumSelectedShapes2D().ToList();
            var shapeids = shapes_2d.Select(s => s.ID).ToList();

            var application = this.Session.Application;
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

            var shapes_2d = this.Session.Selection.EnumSelectedShapes2D().ToList();
            var shapeids = shapes_2d.Select(s => s.ID).ToList();

            var application = this.Session.Application;
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
            var application = this.Session.Application;
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

        public void Send(ShapeSendDirection dir)
        {

            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var selection = Session.Selection.GetSelection();


            if (dir == ShapeSendDirection.ToBack)
            {
                selection.SendToBack();
            }
            else if (dir == ShapeSendDirection.Backward)
            {
                selection.SendBackward();
            }
            else if (dir == ShapeSendDirection.Forward)
            {
                selection.BringForward();
            }
            else if (dir == ShapeSendDirection.ToFront)
            {
                selection.BringToFront();
            }
            else
            {
                throw new System.ArgumentOutOfRangeException("dir");
            }
        }

        public void Align(VA.Drawing.AlignmentHorizontal align)
        {
            if (!this.Session.HasSelectedShapes(2))
            {
                return;
            }

            var cmd = MiscScriptingUtil.halign_to_cmd[align];

            var application = this.Session.Application;
            using (var undoscope = application.CreateUndoScope())
            {
                bool glue_to_guide = false;
                var selection = Session.Selection.GetSelection();
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

            var selection = Session.Selection.GetSelection();
            var shapeids = selection.GetIDs();

            var application = this.Session.Application;
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

            var application = this.Session.Application;
            using (var undoscope = application.CreateUndoScope())
            {
                bool glue_to_guide = false;
                var selection = Session.Selection.GetSelection();
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

            var selection = Session.Selection.GetSelection();
            var shapeids = selection.GetIDs();

            var application = this.Session.Application;
            using (var undoscope = application.CreateUndoScope())
            {
                VA.Layout.LayoutHelper.AlignTo(application.ActivePage, shapeids,align,y);
            }
        }

        public int GetSelectedShapeCount()
        {
            if (!this.Session.HasActiveDrawing())
            {
                return 0;
            }

            var selection = Session.Selection.GetSelection();
            return selection.Count;
        }

        public IList<VA.Layout.XFormCells> GetXForm()
        {
            if (!this.Session.HasSelectedShapes())
            {
                return new List<VA.Layout.XFormCells>(0);
            }

            var selection = Session.Selection.GetSelection();
            var shapeids = selection.GetIDs();

            var page = this.Session.Application.ActivePage;
            var data = VA.Layout.LayoutHelper.GetXForm(page, shapeids);
            return data;
        }

        public IVisio.Shape Group()
        {
            if (!this.Session.HasSelectedShapes(2))
            {
                return null;
            }

            var selection = this.Session.Selection.GetSelection();
            var g = selection.Group();
            return g;
        }

        public void Ungroup()
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            this.Session.Application.DoCmd((short)IVisio.VisUICmds.visCmdObjectUngroup);
        }

        private static VA.ShapeSheet.SRC src_LockAspect = VA.ShapeSheet.SRCConstants.LockAspect;
        private static VA.ShapeSheet.SRC src_LockBegin = VA.ShapeSheet.SRCConstants.LockBegin;
        private static VA.ShapeSheet.SRC src_LockCalcWH = VA.ShapeSheet.SRCConstants.LockCalcWH;
        private static VA.ShapeSheet.SRC src_LockCrop = VA.ShapeSheet.SRCConstants.LockCrop;
        private static VA.ShapeSheet.SRC src_LockCustProp = VA.ShapeSheet.SRCConstants.LockCustProp;
        private static VA.ShapeSheet.SRC src_LockDelete = VA.ShapeSheet.SRCConstants.LockDelete;
        private static VA.ShapeSheet.SRC src_LockEnd = VA.ShapeSheet.SRCConstants.LockEnd;
        private static VA.ShapeSheet.SRC src_LockFormat = VA.ShapeSheet.SRCConstants.LockFormat;
        private static VA.ShapeSheet.SRC src_LockFromGroupFormat = VA.ShapeSheet.SRCConstants.LockFromGroupFormat;
        private static VA.ShapeSheet.SRC src_LockGroup = VA.ShapeSheet.SRCConstants.LockGroup;
        private static VA.ShapeSheet.SRC src_LockHeight = VA.ShapeSheet.SRCConstants.LockHeight;
        private static VA.ShapeSheet.SRC src_LockMoveX = VA.ShapeSheet.SRCConstants.LockMoveX;
        private static VA.ShapeSheet.SRC src_LockMoveY = VA.ShapeSheet.SRCConstants.LockMoveY;
        private static VA.ShapeSheet.SRC src_LockRotate = VA.ShapeSheet.SRCConstants.LockRotate;
        private static VA.ShapeSheet.SRC src_LockSelect = VA.ShapeSheet.SRCConstants.LockSelect;
        private static VA.ShapeSheet.SRC src_LockTextEdit = VA.ShapeSheet.SRCConstants.LockTextEdit;
        private static VA.ShapeSheet.SRC src_LockThemeColors = VA.ShapeSheet.SRCConstants.LockThemeColors;
        private static VA.ShapeSheet.SRC src_LockThemeEffects = VA.ShapeSheet.SRCConstants.LockThemeEffects;
        private static VA.ShapeSheet.SRC src_LockVtxEdit = VA.ShapeSheet.SRCConstants.LockVtxEdit;
        private static VA.ShapeSheet.SRC src_LockWidth = VA.ShapeSheet.SRCConstants.LockWidth;
        private static VA.ShapeSheet.SRC[] lockcells = new[]
                                                 {
                                                     src_LockAspect,
                                                     src_LockBegin,
                                                     src_LockCalcWH,
                                                     src_LockCrop,
                                                     src_LockCustProp,
                                                     src_LockDelete,
                                                     src_LockEnd,
                                                     src_LockFormat,
                                                     src_LockFromGroupFormat,
                                                     src_LockGroup,
                                                     src_LockHeight,
                                                     src_LockMoveX,
                                                     src_LockMoveY,
                                                     src_LockRotate,
                                                     src_LockSelect,
                                                     src_LockTextEdit,
                                                     src_LockThemeColors,
                                                     src_LockThemeEffects,
                                                     src_LockVtxEdit,
                                                     src_LockWidth

                                                 };

        private void SetLockCells(VA.ShapeSheet.SRC[] srcs, double val)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }
            var invariant_culture = System.Globalization.CultureInfo.InvariantCulture;
            var formula = val.ToString(invariant_culture);
            var formulas = srcs.Select(src => formula).ToList();
            IVisio.VisGetSetArgs flags = 0;
            this.Session.ShapeSheet.SetFormula(lockcells, formulas, flags);
        }

        public void LockAll()
        {
            SetLockCells(lockcells, 1.0);
        }

        public void UnlockAll()
        {
            SetLockCells(lockcells, 0.0);
        }

        public void SetWidth(double w)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var selection = this.Session.Selection.GetSelection();
            var shapeids = selection.GetIDs();
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            foreach (int shapeid in shapeids)
            {
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.Width, w);
            }

            var application = this.Session.Application;
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

            var selection = this.Session.Selection.GetSelection();
            var shapeids = selection.GetIDs();
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            foreach (int shapeid in shapeids)
            {
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.Height, h);
            }

            var application = this.Session.Application;
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

            var selection = this.Session.Selection.GetSelection();
            var shapeids = selection.GetIDs();
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            foreach (int shapeid in shapeids)
            {
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.Width, w);
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.Height, h);
            }

            var application = this.Session.Application;
            using (var undoscope = application.CreateUndoScope())
            {
                var active_page = application.ActivePage;
                update.Execute(active_page);
            }
        }
    }
}
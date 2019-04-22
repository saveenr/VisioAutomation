using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class ArrangeCommands : CommandSet
    {
        internal ArrangeCommands(Client client) :
            base(client)
        {

        }

        public void Nudge(TargetSelection targetselection, double dx, double dy)
        {
            if (dx == 0.0 && dy == 0.0)
            {
                return;
            }

            targetselection = targetselection.Resolve(this._client);

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(Nudge)))
            {
                var unitcode = IVisio.VisUnitCodes.visInches;

                // Move method: http://msdn.microsoft.com/en-us/library/ms367549.aspx   
                targetselection.Selection.Move(dx, dy, unitcode);
            }
        }

        private static void _send_selection(IVisio.Selection selection, Models.ShapeSendDirection dir)
        {

            if (dir == Models.ShapeSendDirection.ToBack)
            {
                selection.SendToBack();
            }
            else if (dir == Models.ShapeSendDirection.Backward)
            {
                selection.SendBackward();
            }
            else if (dir == Models.ShapeSendDirection.Forward)
            {
                selection.BringForward();
            }
            else if (dir == Models.ShapeSendDirection.ToFront)
            {
                selection.BringToFront();
            }
        }


        public void Send(Models.ShapeSendDirection dir)
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.RequireDocument);

            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            ArrangeCommands._send_selection(selection, dir);
        }

        public void AlignHorizontal(TargetSelection targetselection, Models.AlignmentHorizontal align)
        {
            targetselection = targetselection.Resolve(this._client);

            IVisio.VisHorizontalAlignTypes halign;
            var valign = IVisio.VisVerticalAlignTypes.visVertAlignNone;

            switch (align)
            {
                case VisioScripting.Models.AlignmentHorizontal.Left:
                    halign = IVisio.VisHorizontalAlignTypes.visHorzAlignLeft;
                    break;
                case VisioScripting.Models.AlignmentHorizontal.Center:
                    halign = IVisio.VisHorizontalAlignTypes.visHorzAlignCenter;
                    break;
                case VisioScripting.Models.AlignmentHorizontal.Right:
                    halign = IVisio.VisHorizontalAlignTypes.visHorzAlignRight;
                    break;
                default: throw new System.ArgumentOutOfRangeException();
            }

            const bool glue_to_guide = false;

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(AlignHorizontal)))
            {
                targetselection.Selection.Align(halign, valign, glue_to_guide);
            }
        }

        public void AlignVertical(TargetSelection targetselection, Models.AlignmentVertical align)
        {
            targetselection = targetselection.Resolve(this._client);

            // Set the align enums
            var halign = IVisio.VisHorizontalAlignTypes.visHorzAlignNone;
            IVisio.VisVerticalAlignTypes valign;
            switch (align)
            {
                case VisioScripting.Models.AlignmentVertical.Top:
                    valign = IVisio.VisVerticalAlignTypes.visVertAlignTop;
                    break;
                case VisioScripting.Models.AlignmentVertical.Center:
                    valign = IVisio.VisVerticalAlignTypes.visVertAlignMiddle;
                    break;
                case VisioScripting.Models.AlignmentVertical.Bottom:
                    valign = IVisio.VisVerticalAlignTypes.visVertAlignBottom;
                    break;
                default: throw new System.ArgumentOutOfRangeException();
            }

            const bool glue_to_guide = false;

            // Perform the alignment
            using (var undoscope = this._client.Undo.NewUndoScope(nameof(AlignVertical)))
            {
                targetselection.Selection.Align(halign, valign, glue_to_guide);
            }
        }

        public void DistributenOnAxis(TargetShapes targetshapes, Models.Axis axis, double spacing)
        {
            targetshapes = targetshapes.Resolve(this._client);

            if (targetshapes.Shapes.Count < 1)
            {
                return;
            }

            var page = targetshapes.Shapes[0].ContainingPage;
            var targetshapeids = targetshapes.ToShapeIDs();

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DistributeOnAxis)))
            {
                VisioScripting.Helpers.ArrangeHelper._distribute_with_spacing(page, targetshapeids, axis, spacing);
            }
        }

        public void DistributeOnAxis(VisioScripting.TargetSelection targetselection, Models.Axis axis)
        {
            targetselection = targetselection.Resolve(this._client);

            if (targetselection.Selection.Count < 2)
            {
                return;
            }

            IVisio.VisUICmds cmd;

            switch (axis)
            {
                case VisioScripting.Models.Axis.XAxis:
                    cmd = IVisio.VisUICmds.visCmdDistributeHSpace;
                    break;
                case VisioScripting.Models.Axis.YAxis:
                    cmd = IVisio.VisUICmds.visCmdDistributeVSpace;
                    break;
                default:
                    throw new System.ArgumentOutOfRangeException();
            }

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DistributeOnAxis)))
            {
                targetselection.Selection.Application.DoCmd((short) cmd);
            }
        }

        public void DistributeHorizontal(TargetSelection targetselection, Models.AlignmentHorizontal halign)
        {
            targetselection = targetselection.Resolve(this._client);

            if (targetselection.Selection.Count < 2)
            {
                return;
            }

            IVisio.VisUICmds cmd;

            switch (halign)
            {
                case VisioScripting.Models.AlignmentHorizontal.Left:
                    cmd = IVisio.VisUICmds.visCmdDistributeLeft;
                    break;
                case VisioScripting.Models.AlignmentHorizontal.Center:
                    cmd = IVisio.VisUICmds.visCmdDistributeCenter;
                    break;
                case VisioScripting.Models.AlignmentHorizontal.Right:
                    cmd = IVisio.VisUICmds.visCmdDistributeRight;
                    break;
                default: throw new System.ArgumentOutOfRangeException();
            }

            var app = targetselection.Selection.Application;
            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DistributeHorizontal)))
            {
                app.DoCmd((short)cmd);
            }
        }

        public void DistributeVertical(TargetSelection targetselection, Models.AlignmentVertical valign)
        {
            targetselection = targetselection.Resolve(this._client);

            if (targetselection.Selection.Count < 2)
            {
                return;
            }


            IVisio.VisUICmds cmd;
            switch (valign)
            {
                case VisioScripting.Models.AlignmentVertical.Top:
                    cmd = IVisio.VisUICmds.visCmdDistributeTop;
                    break;
                case VisioScripting.Models.AlignmentVertical.Center:
                    cmd = IVisio.VisUICmds.visCmdDistributeMiddle;
                    break;
                case VisioScripting.Models.AlignmentVertical.Bottom:
                    cmd = IVisio.VisUICmds.visCmdDistributeBottom;
                    break;
                default: throw new System.ArgumentOutOfRangeException();
            }


            var app = targetselection.Selection.Application;
            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DistributeVertical)))
            {
                app.DoCmd((short)cmd);
            }
        }
    }
}
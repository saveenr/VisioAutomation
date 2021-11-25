

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

            targetselection = targetselection.ResolveToSelection(this._client);

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
            targetselection = targetselection.ResolveToSelection(this._client);
            var valign = IVisio.VisVerticalAlignTypes.visVertAlignNone;

            var halign = align switch
            {
                VisioScripting.Models.AlignmentHorizontal.Left => IVisio.VisHorizontalAlignTypes.visHorzAlignLeft,
                VisioScripting.Models.AlignmentHorizontal.Center => IVisio.VisHorizontalAlignTypes.visHorzAlignCenter,
                VisioScripting.Models.AlignmentHorizontal.Right => IVisio.VisHorizontalAlignTypes.visHorzAlignRight,
                _ => throw new System.ArgumentOutOfRangeException(),
            };

            const bool glue_to_guide = false;

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(AlignHorizontal)))
            {
                targetselection.Selection.Align(halign, valign, glue_to_guide);
            }
        }

        public void AlignVertical(TargetSelection targetselection, Models.AlignmentVertical align)
        {
            targetselection = targetselection.ResolveToSelection(this._client);
            var halign = IVisio.VisHorizontalAlignTypes.visHorzAlignNone;
            var valign = align switch
            {
                VisioScripting.Models.AlignmentVertical.Top => IVisio.VisVerticalAlignTypes.visVertAlignTop,
                VisioScripting.Models.AlignmentVertical.Center => IVisio.VisVerticalAlignTypes.visVertAlignMiddle,
                VisioScripting.Models.AlignmentVertical.Bottom => IVisio.VisVerticalAlignTypes.visVertAlignBottom,
                _ => throw new System.ArgumentOutOfRangeException(),
            };
            const bool glue_to_guide = false;

            // Perform the alignment
            using (var undoscope = this._client.Undo.NewUndoScope(nameof(AlignVertical)))
            {
                targetselection.Selection.Align(halign, valign, glue_to_guide);
            }
        }

        public void DistributeOnAxis(VisioScripting.TargetSelection targetselection, Models.Axis axis)
        {
            targetselection = targetselection.ResolveToSelection(this._client);

            if (targetselection.Selection.Count < 2)
            {
                return;
            }

            var cmd = axis switch
            {
                VisioScripting.Models.Axis.XAxis => IVisio.VisUICmds.visCmdDistributeHSpace,
                VisioScripting.Models.Axis.YAxis => IVisio.VisUICmds.visCmdDistributeVSpace,
                _ => throw new System.ArgumentOutOfRangeException(),
            };
            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DistributeOnAxis)))
            {
                targetselection.Selection.Application.DoCmd((short) cmd);
            }
        }

        public void DistributeHorizontal(TargetSelection targetselection, Models.AlignmentHorizontal halign)
        {
            targetselection = targetselection.ResolveToSelection(this._client);

            if (targetselection.Selection.Count < 2)
            {
                return;
            }

            var cmd = halign switch
            {
                VisioScripting.Models.AlignmentHorizontal.Left => IVisio.VisUICmds.visCmdDistributeLeft,
                VisioScripting.Models.AlignmentHorizontal.Center => IVisio.VisUICmds.visCmdDistributeCenter,
                VisioScripting.Models.AlignmentHorizontal.Right => IVisio.VisUICmds.visCmdDistributeRight,
                _ => throw new System.ArgumentOutOfRangeException(),
            };
            var app = targetselection.Selection.Application;
            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DistributeHorizontal)))
            {
                app.DoCmd((short)cmd);
            }
        }

        public void DistributeVertical(TargetSelection targetselection, Models.AlignmentVertical valign)
        {
            targetselection = targetselection.ResolveToSelection(this._client);

            if (targetselection.Selection.Count < 2)
            {
                return;
            }
            
            var cmd = valign switch
            {
                VisioScripting.Models.AlignmentVertical.Top => IVisio.VisUICmds.visCmdDistributeTop,
                VisioScripting.Models.AlignmentVertical.Center => IVisio.VisUICmds.visCmdDistributeMiddle,
                VisioScripting.Models.AlignmentVertical.Bottom => IVisio.VisUICmds.visCmdDistributeBottom,
                _ => throw new System.ArgumentOutOfRangeException(),
            };

            var app = targetselection.Selection.Application;
            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DistributeVertical)))
            {
                app.DoCmd((short)cmd);
            }
        }
    }
}
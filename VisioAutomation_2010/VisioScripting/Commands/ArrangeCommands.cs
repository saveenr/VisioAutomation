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

            var cmdtarget = this._client.GetCommandTargetDocument();

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(Nudge)))
            {
                var window = cmdtarget.Application.ActiveWindow;
                var selection = window.Selection;
                var unitcode = IVisio.VisUnitCodes.visInches;

                // Move method: http://msdn.microsoft.com/en-us/library/ms367549.aspx   
                selection.Move(dx, dy, unitcode);
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
            var cmdtarget = this._client.GetCommandTargetDocument();
            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            ArrangeCommands._send_selection(selection, dir);
        }

        public void AlignHorizontal(TargetSelection targetselection, Models.AlignmentHorizontal align)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();

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
                var window = cmdtarget.Application.ActiveWindow;
                var selection = window.Selection;
                selection.Align(halign, valign, glue_to_guide);
            }
        }

        public void AlignVertical(TargetSelection targetselection, Models.AlignmentVertical align)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();

            // Set the align enums
            var halign = IVisio.VisHorizontalAlignTypes.visHorzAlignNone;
            IVisio.VisVerticalAlignTypes valign;
            switch (align)
            {
                case VisioScripting.Models.AlignmentVertical.Top: valign = IVisio.VisVerticalAlignTypes.visVertAlignTop; break;
                case VisioScripting.Models.AlignmentVertical.Center: valign = IVisio.VisVerticalAlignTypes.visVertAlignMiddle; break;
                case VisioScripting.Models.AlignmentVertical.Bottom: valign = IVisio.VisVerticalAlignTypes.visVertAlignBottom; break;
                default: throw new System.ArgumentOutOfRangeException();
            }

            const bool glue_to_guide = false;

            // Perform the alignment
            using (var undoscope = this._client.Undo.NewUndoScope(nameof(AlignVertical)))
            {
                var window = cmdtarget.Application.ActiveWindow;
                var selection = window.Selection;
                selection.Align(halign, valign, glue_to_guide);
            }
        }

    }
}
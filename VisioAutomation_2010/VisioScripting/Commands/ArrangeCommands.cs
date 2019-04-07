using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class ArrangeCommands : CommandSet
    {
        internal ArrangeCommands(Client client) :
            base(client)
        {

        }

        public void NudgeSelection(TargetSelection targetselection, double dx, double dy)
        {
            if (dx == 0.0 && dy == 0.0)
            {
                return;
            }

            var cmdtarget = this._client.GetCommandTargetDocument();

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(NudgeSelection)))
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


        public void SendSelection(Models.ShapeSendDirection dir)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();
            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            ArrangeCommands._send_selection(selection, dir);
        }
    }
}
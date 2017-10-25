using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class AlignCommands : CommandSet
    {
        internal AlignCommands(Client client) :
            base(client)
        {

        }

        public void AlignSelectionHorizontal(Models.AlignmentHorizontal align)
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

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(AlignSelectionHorizontal)))
            {
                var window = cmdtarget.Application.ActiveWindow;
                var selection = window.Selection;
                selection.Align(halign, valign, glue_to_guide);
            }
        }

        public void AlignSelectionVertical(Models.AlignmentVertical align)
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
            using (var undoscope = this._client.Undo.NewUndoScope(nameof(AlignSelectionVertical)))
            {
                var window = cmdtarget.Application.ActiveWindow;
                var selection = window.Selection;
                selection.Align(halign, valign, glue_to_guide);
            }
        }

    }
}
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class AlignCommands : CommandSet
    {
        internal AlignCommands(Client client) :
            base(client)
        {

        }

        public void AlignHorizontal(VisioScripting.Models.TargetShapes targets, VisioScripting.Models.AlignmentHorizontal align)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 2)
            {
                return;
            }

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

            using (var undoscope = this._client.Application.NewUndoScope("Align Horizontal"))
            {
                var window = cmdtarget.Application.ActiveWindow;
                var selection = window.Selection;
                selection.Align(halign, valign, glue_to_guide);
            }
        }

        public void AlignVertical(VisioScripting.Models.TargetShapes targets, VisioScripting.Models.AlignmentVertical align)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 2)
            {
                return;
            }

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
            using (var undoscope = this._client.Application.NewUndoScope("Align Vertical"))
            {
                var window = cmdtarget.Application.ActiveWindow;
                var selection = window.Selection;
                selection.Align(halign, valign, glue_to_guide);
            }
        }

    }
}
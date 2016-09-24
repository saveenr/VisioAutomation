using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Commands
{
    public class AlignCommands : CommandSet
    {
        internal AlignCommands(Client client) :
            base(client)
        {

        }

        public void AlignHorizontal(TargetShapes targets, VisioAutomation.Drawing.Layout.AlignmentHorizontal align)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 2)
            {
                return;
            }

            IVisio.VisHorizontalAlignTypes halign;
            var valign = IVisio.VisVerticalAlignTypes.visVertAlignNone;

            switch (align)
            {
                case VisioAutomation.Drawing.Layout.AlignmentHorizontal.Left:
                    halign = IVisio.VisHorizontalAlignTypes.visHorzAlignLeft;
                    break;
                case VisioAutomation.Drawing.Layout.AlignmentHorizontal.Center:
                    halign = IVisio.VisHorizontalAlignTypes.visHorzAlignCenter;
                    break;
                case VisioAutomation.Drawing.Layout.AlignmentHorizontal.Right:
                    halign = IVisio.VisHorizontalAlignTypes.visHorzAlignRight;
                    break;
                default: throw new System.ArgumentOutOfRangeException();
            }

            const bool glue_to_guide = false;

            using (var undoscope = this._client.Application.NewUndoScope("Align Shapes"))
            {
                var selection = this._client.Selection.Get();
                selection.Align(halign, valign, glue_to_guide);
            }
        }

        public void AlignVertical(TargetShapes targets, VisioAutomation.Drawing.Layout.AlignmentVertical align)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

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
                case VisioAutomation.Drawing.Layout.AlignmentVertical.Top: valign = IVisio.VisVerticalAlignTypes.visVertAlignTop; break;
                case VisioAutomation.Drawing.Layout.AlignmentVertical.Center: valign = IVisio.VisVerticalAlignTypes.visVertAlignMiddle; break;
                case VisioAutomation.Drawing.Layout.AlignmentVertical.Bottom: valign = IVisio.VisVerticalAlignTypes.visVertAlignBottom; break;
                default: throw new System.ArgumentOutOfRangeException();
            }

            const bool glue_to_guide = false;

            // Perform the alignment
            using (var undoscope = this._client.Application.NewUndoScope("Align Shapes"))
            {
                var selection = this._client.Selection.Get();
                selection.Align(halign, valign, glue_to_guide);
            }
        }

    }
}
using VisioAutomation.Drawing.Layout;

namespace VisioAutomation.Scripting.Commands
{
    public class AlignCommands : CommandSet
    {
        internal AlignCommands(Client client) :
            base(client)
        {

        }

        public void AlignHorizontal(TargetShapes targets, AlignmentHorizontal align)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 2)
            {
                return;
            }

            Microsoft.Office.Interop.Visio.VisHorizontalAlignTypes halign;
            var valign = Microsoft.Office.Interop.Visio.VisVerticalAlignTypes.visVertAlignNone;

            switch (align)
            {
                case AlignmentHorizontal.Left:
                    halign = Microsoft.Office.Interop.Visio.VisHorizontalAlignTypes.visHorzAlignLeft;
                    break;
                case AlignmentHorizontal.Center:
                    halign = Microsoft.Office.Interop.Visio.VisHorizontalAlignTypes.visHorzAlignCenter;
                    break;
                case AlignmentHorizontal.Right:
                    halign = Microsoft.Office.Interop.Visio.VisHorizontalAlignTypes.visHorzAlignRight;
                    break;
                default: throw new System.ArgumentOutOfRangeException();
            }

            const bool glue_to_guide = false;

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Align Shapes"))
            {
                var selection = this._client.Selection.Get();
                selection.Align(halign, valign, glue_to_guide);
            }
        }

        public void AlignVertical(TargetShapes targets, AlignmentVertical align)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 2)
            {
                return;
            }

            // Set the align enums
            var halign = Microsoft.Office.Interop.Visio.VisHorizontalAlignTypes.visHorzAlignNone;
            Microsoft.Office.Interop.Visio.VisVerticalAlignTypes valign;
            switch (align)
            {
                case AlignmentVertical.Top: valign = Microsoft.Office.Interop.Visio.VisVerticalAlignTypes.visVertAlignTop; break;
                case AlignmentVertical.Center: valign = Microsoft.Office.Interop.Visio.VisVerticalAlignTypes.visVertAlignMiddle; break;
                case AlignmentVertical.Bottom: valign = Microsoft.Office.Interop.Visio.VisVerticalAlignTypes.visVertAlignBottom; break;
                default: throw new System.ArgumentOutOfRangeException();
            }

            const bool glue_to_guide = false;

            // Perform the alignment
            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Align Shapes"))
            {
                var selection = this._client.Selection.Get();
                selection.Align(halign, valign, glue_to_guide);
            }
        }

    }
}
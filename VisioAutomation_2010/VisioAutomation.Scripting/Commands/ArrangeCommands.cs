using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Shapes.Locking;

namespace VisioAutomation.Scripting.Commands
{
    public class ArrangeCommands : CommandSet
    {
        internal ArrangeCommands(Client client) :
            base(client)
        {

        }

        public void Nudge(TargetShapes targets, double dx, double dy)
        {
            if (dx == 0.0 && dy == 0.0)
            {
                return;
            }

            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Application.NewUndoScope("Nudge"))
            {
                var selection = this._client.Selection.Get();
                var unitcode = IVisio.VisUnitCodes.visInches;

                // Move method: http://msdn.microsoft.com/en-us/library/ms367549.aspx   
                selection.Move(dx, dy, unitcode);
            }
        }


        public void Send(TargetShapes targets, Selections.ShapeSendDirection dir)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 1)
            {
                return;
            }

            var selection = this._client.Selection.Get();
            Selections.SelectionHelper.SendShapes(selection, dir);
        }

        public void SetLock(TargetShapes targets, LockCells lockcells)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            targets = targets.ResolveShapes(this._client);
            if (targets.Shapes.Count < 1)
            {
                return;
            }

            var page = this._client.Page.Get();
            var target_shapeids = targets.ToShapeIDs();
            var writer = new VisioAutomation.ShapeSheet.ShapeSheetWriter();

            foreach (int shapeid in target_shapeids.ShapeIDs)
            {
                lockcells.SetFormulas((short)shapeid, writer);
            }

            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(page);

            using (var undoscope = this._client.Application.NewUndoScope("Set Lock Properties"))
            {
                writer.Commit(surface);
            }
        }

        public void SetSize(TargetShapes targets, double? w, double? h)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            targets = targets.ResolveShapes(this._client);
            if (targets.Shapes.Count < 1)
            {
                return;
            }

            var active_page = this._client.Page.Get();
            var shapeids = targets.ToShapeIDs();
            var writer = new VisioAutomation.ShapeSheet.ShapeSheetWriter();
            foreach (int shapeid in shapeids.ShapeIDs)
            {
                if (w.HasValue && w.Value>=0)
                {
                    writer.SetFormula((short)shapeid, VisioAutomation.ShapeSheet.SRCConstants.Width, w.Value);
                }
                if (h.HasValue && h.Value >= 0)
                {
                    writer.SetFormula((short)shapeid, VisioAutomation.ShapeSheet.SRCConstants.Height, h.Value);                    
                }
            }

            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(active_page);

            using (var undoscope = this._client.Application.NewUndoScope("Set Shape Size"))
            {
                writer.Commit(surface);
            }
        }
    }
}
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using VisioAutomation.Drawing.Layout;
using VisioAutomation.ShapeSheet.Writers;

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

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Nudge Shapes"))
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

        public void SetLock(TargetShapes targets, Shapes.LockCells lockcells)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapes(this._client);
            if (shapes.Count < 1)
            {
                return;
            } 

            var selection = this._client.Selection.Get();
            var shapeids = selection.GetIDs();
            var writer = new FormulaWriterSIDSRC();

            foreach (int shapeid in shapeids)
            {
                lockcells.SetFormulas((short)shapeid, writer);
            }

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Set Shape Lock Properties"))
            {
                var active_page = application.ActivePage;
                writer.Commit(active_page);
            }
        }

        public void SetSize(TargetShapes targets, double? w, double? h)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapes(this._client);
            if (shapes.Count < 1)
            {
                return;
            } 

            var shapeids = shapes.Select(s=>s.ID).ToList();
            var writer = new FormulaWriterSIDSRC();
            foreach (int shapeid in shapeids)
            {
                if (w.HasValue && w.Value>=0)
                {
                    writer.SetFormula((short)shapeid, ShapeSheet.SRCConstants.Width, w.Value);
                }
                if (h.HasValue && h.Value >= 0)
                {
                    writer.SetFormula((short)shapeid, ShapeSheet.SRCConstants.Height, h.Value);                    
                }
            }

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Set Shape Size"))
            {
                var active_page = application.ActivePage;
                writer.Commit(active_page);
            }
        }
    }
}
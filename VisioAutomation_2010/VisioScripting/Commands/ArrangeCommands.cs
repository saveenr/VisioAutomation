using System.Collections.Generic;
using VisioAutomation.Shapes;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Writers;

namespace VisioScripting.Commands
{
    public class ArrangeCommands : CommandSet
    {
        internal ArrangeCommands(Client client) :
            base(client)
        {

        }

        public void Nudge(VisioScripting.Models.TargetShapes targets, double dx, double dy)
        {
            if (dx == 0.0 && dy == 0.0)
            {
                return;
            }

            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Application.NewUndoScope("Nudge"))
            {
                var window = cmdtarget.Application.ActiveWindow;
                var selection = window.Selection;
                var unitcode = Microsoft.Office.Interop.Visio.VisUnitCodes.visInches;

                // Move method: http://msdn.microsoft.com/en-us/library/ms367549.aspx   
                selection.Move(dx, dy, unitcode);
            }
        }

        private static void SendShapes(Microsoft.Office.Interop.Visio.Selection selection, Models.ShapeSendDirection dir)
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


        public void Send(VisioScripting.Models.TargetShapes targets, Models.ShapeSendDirection dir)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 1)
            {
                return;
            }

            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            ArrangeCommands.SendShapes(selection, dir);
        }

        public void SetLock(VisioScripting.Models.TargetShapes targets, LockCells lockcells)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            targets = targets.ResolveShapes(this._client);
            if (targets.Shapes.Count < 1)
            {
                return;
            }

            var page = cmdtarget.ActivePage;
            var target_shapeids = targets.ToShapeIDs();
            var writer = new SidSrcWriter();

            foreach (int shapeid in target_shapeids.ShapeIDs)
            {
                lockcells.SetFormulas(writer, (short)shapeid);
            }

            using (var undoscope = this._client.Application.NewUndoScope("Set Lock Properties"))
            {
                writer.Commit(page);
            }
        }


        public Dictionary<int,LockCells> GetLock(VisioScripting.Models.TargetShapes targets)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            targets = targets.ResolveShapes(this._client);
            if (targets.Shapes.Count < 1)
            {
                return new Dictionary<int, LockCells>();
            }

            var dic = new Dictionary<int, LockCells>();

            var page = cmdtarget.ActivePage;
            var target_shapeids = targets.ToShapeIDs();

            var cells = VisioAutomation.Shapes.LockCells.GetCells(page, target_shapeids.ShapeIDs, CellValueType.Formula);

            for (int i = 0; i < target_shapeids.ShapeIDs.Count; i++)
            {
                var shapeid = target_shapeids.ShapeIDs[i];
                var cur_cells = cells[i];
                dic[shapeid] = cur_cells;
            }

            return dic;
        }

        public void SetSize(VisioScripting.Models.TargetShapes targets, double? w, double? h)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            targets = targets.ResolveShapes(this._client);
            if (targets.Shapes.Count < 1)
            {
                return;
            }

            var active_page = cmdtarget.ActivePage;
            var shapeids = targets.ToShapeIDs();
            var writer = new SidSrcWriter();
            foreach (int shapeid in shapeids.ShapeIDs)
            {
                if (w.HasValue && w.Value>=0)
                {
                    writer.SetFormula((short)shapeid, VisioAutomation.ShapeSheet.SrcConstants.XFormWidth, w.Value);
                }
                if (h.HasValue && h.Value >= 0)
                {
                    writer.SetFormula((short)shapeid, VisioAutomation.ShapeSheet.SrcConstants.XFormHeight, h.Value);                    
                }
            }

            using (var undoscope = this._client.Application.NewUndoScope("Set Shape Size"))
            {
                writer.Commit(active_page);
            }
        }
    }
}
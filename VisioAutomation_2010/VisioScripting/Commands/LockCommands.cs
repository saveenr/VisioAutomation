using System.Collections.Generic;
using VisioAutomation.Shapes;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Writers;

namespace VisioScripting.Commands
{
    public class LockCommands : CommandSet
    {
        internal LockCommands(Client client) :
            base(client)
        {

        }

        public void SetLockCells(VisioScripting.Models.TargetShapes targets, LockCells lockcells)
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

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


        public Dictionary<int,LockCells> GetLockCells(VisioScripting.Models.TargetShapes targets, CellValueType cvt)
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            targets = targets.ResolveShapes(this._client);
            if (targets.Shapes.Count < 1)
            {
                return new Dictionary<int, LockCells>();
            }

            var dic = new Dictionary<int, LockCells>();

            var page = cmdtarget.ActivePage;
            var target_shapeids = targets.ToShapeIDs();

            var cells = VisioAutomation.Shapes.LockCells.GetCells(page, target_shapeids.ShapeIDs, cvt);

            for (int i = 0; i < target_shapeids.ShapeIDs.Count; i++)
            {
                var shapeid = target_shapeids.ShapeIDs[i];
                var cur_cells = cells[i];
                dic[shapeid] = cur_cells;
            }

            return dic;
        }
    }
}
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Shapes;
using VASS=VisioAutomation.ShapeSheet;

namespace VisioScripting.Commands
{
    public class LockCommands : CommandSet
    {
        internal LockCommands(Client client) :
            base(client)
        {

        }

        public void SetLockCells(Models.TargetShapes targets, LockCells lockcells)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            targets = targets.ResolveShapes(this._client);
            if (targets.Shapes.Count < 1)
            {
                return;
            }

            var page = cmdtarget.ActivePage;
            var target_shapeids = targets.ToShapeIDs();
            var writer = new VASS.Writers.SidSrcWriter();

            foreach (int shapeid in target_shapeids.ShapeIDs)
            {
                writer.SetValues((short)shapeid, lockcells);
            }

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(SetLockCells)))
            {
                writer.Commit(page, VASS.CellValueType.Formula);
            }
        }


        public Dictionary<int,LockCells> GetLockCells(Models.TargetShapes targets, VASS.CellValueType cvt)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            targets = targets.ResolveShapes(this._client);
            if (targets.Shapes.Count < 1)
            {
                return new Dictionary<int, LockCells>();
            }

            var dic = new Dictionary<int, LockCells>();

            var page = cmdtarget.ActivePage;
            var target_shapeids = targets.Shapes.Select(s => (int)s.ID16).ToList();

            var cells = VisioAutomation.Shapes.LockCells.GetCells(page, target_shapeids, cvt);

            for (int i = 0; i < target_shapeids.Count; i++)
            {
                var shapeid = target_shapeids[i];
                var cur_cells = cells[i];
                dic[shapeid] = cur_cells;
            }

            return dic;
        }
    }
}
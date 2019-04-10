using System.Collections.Generic;
using VA=VisioAutomation;
using VASS=VisioAutomation.ShapeSheet;

namespace VisioScripting.Commands
{
    public class LockCommands : CommandSet
    {
        internal LockCommands(Client client) :
            base(client)
        {

        }

        public void SetLockCells(TargetShapes targetshapes, VA.Shapes.LockCells lockcells)
        {
            targetshapes = targetshapes.Resolve(this._client);
            if (targetshapes.Shapes.Count < 1)
            {
                return;
            }

            var page = targetshapes.Shapes[0].ContainingPage;
            var targetshapeids = targetshapes.ToShapeIDs();
            var writer = new VASS.Writers.SidSrcWriter();

            foreach (int shapeid in targetshapeids)
            {
                writer.SetValues((short)shapeid, lockcells);
            }

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(SetLockCells)))
            {
                writer.Commit(page, VASS.CellValueType.Formula);
            }
        }


        public Dictionary<int, VA.Shapes.LockCells> GetLockCells(TargetShapes targetshapes, VASS.CellValueType cvt)
        {

            targetshapes = targetshapes.Resolve(this._client);
            if (targetshapes.Shapes.Count < 1)
            {
                return new Dictionary<int, VA.Shapes.LockCells>();
            }

            var dic = new Dictionary<int, VA.Shapes.LockCells>();

            var page = targetshapes.Shapes[0].ContainingPage;

            var target_shapeids = targetshapes.ToShapeIDs();

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
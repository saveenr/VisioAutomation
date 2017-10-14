using System.Collections.Generic;
using VisioAutomation.Shapes;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class ControlCommands : CommandSet
    {
        internal ControlCommands(Client client) :
            base(client)
        {

        }

        public List<int> AddControlToShapes(VisioScripting.Models.TargetShapes targets, ControlCells ctrl)
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);

            if (ctrl == null)
            {
                throw new System.ArgumentNullException(nameof(ctrl));
            }

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return new List<int>(0);
            }


            var control_indices = new List<int>();

            using (var undoscope = this._client.Application.NewUndoScope("Add Control"))
            {
                foreach (var shape in targets.Shapes)
                {
                    int ci = ControlHelper.Add(shape, ctrl);
                    control_indices.Add(ci);
                }
            }

            return control_indices;
        }

        public void DeleteControlWithIndex(VisioScripting.Models.TargetShapes targets, int n)
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Application.NewUndoScope("Delete Control"))
            {
                foreach (var shape in targets.Shapes)
                {
                    ControlHelper.Delete(shape, n);
                }
            }
        }

        public Dictionary<IVisio.Shape, IList<ControlCells>> GetControlCells(VisioScripting.Models.TargetShapes targets, CellValueType cvt)
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return new Dictionary<IVisio.Shape, IList<ControlCells>>(0);
            }

            var dic = new Dictionary<IVisio.Shape, IList<ControlCells>>();
            foreach (var shape in targets.Shapes)
            {
                var controls = ControlCells.GetCells(shape, cvt);
                dic[shape] = controls;
            }
            return dic;
        }
    }
}
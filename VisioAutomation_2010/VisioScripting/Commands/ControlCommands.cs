
using VisioAutomation.Shapes;
using VisioAutomation.ShapeSheet;


namespace VisioScripting.Commands
{
    public class ControlCommands : CommandSet
    {
        internal ControlCommands(Client client) :
            base(client)
        {

        }

        public List<int> AddControlToShapes(TargetShapes targetshapes, ControlCells ctrl)
        {
            if (ctrl == null)
            {
                throw new System.ArgumentNullException(nameof(ctrl));
            }

            targetshapes = targetshapes.ResolveToShapes(this._client);

            if (targetshapes.Shapes.Count < 1)
            {
                return new List<int>(0);
            }

            var control_indices = new List<int>();

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(AddControlToShapes)))
            {
                foreach (var shape in targetshapes.Shapes)
                {
                    int ci = ControlHelper.Add(shape, ctrl);
                    control_indices.Add(ci);
                }
            }

            return control_indices;
        }

        public void DeleteControlWithIndex(TargetShapes targetshapes, int index)
        {
            targetshapes = targetshapes.ResolveToShapes(this._client);

            if (targetshapes.Shapes.Count < 1)
            {
                return;
            }

            // restrict the operation to those shapes that actually have enough
            // controls to qualify for deleting 
            var qualified_shapes = targetshapes.Shapes.Where(shape => ControlHelper.GetCount(shape) > index);

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DeleteControlWithIndex)))
            {
                foreach (var shape in qualified_shapes)
                {
                    ControlHelper.Delete(shape, index);
                }
            }
        }

        public Dictionary<IVisio.Shape, IList<ControlCells>> GetControls(TargetShapes targetshapes, CellValueType cvt)
        {
            targetshapes = targetshapes.ResolveToShapes(this._client);

            if (targetshapes.Shapes.Count < 1)
            {
                return new Dictionary<IVisio.Shape, IList<ControlCells>>(0);
            }

            var dic = new Dictionary<IVisio.Shape, IList<ControlCells>>(targetshapes.Shapes.Count);
            foreach (var shape in targetshapes.Shapes)
            {
                var controls = ControlCells.GetCells(shape, cvt);
                dic[shape] = controls;
            }
            return dic;
        }
    }
}
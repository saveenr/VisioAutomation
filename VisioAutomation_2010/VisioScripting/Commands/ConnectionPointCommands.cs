


namespace VisioScripting.Commands
{
    public class ConnectionPointCommands : CommandSet
    {
        internal ConnectionPointCommands(Client client) :
            base(client)
        {

        }

        public IDictionary<IVisio.Shape, IList<VA.Shapes.ConnectionPointCells>> GetConnectionPoints(TargetShapes targetshapes)
        {
            targetshapes = targetshapes.ResolveToShapes(this._client);

            if (targetshapes.Shapes.Count <1)
            {
                return new Dictionary<IVisio.Shape, IList<VA.Shapes.ConnectionPointCells>>();
            }

            var dicof_shape_to_cxnpoint = new Dictionary<IVisio.Shape, IList<VA.Shapes.ConnectionPointCells>>();
            foreach (var shape in targetshapes.Shapes)
            {
                var cp = VisioAutomation.Shapes.ConnectionPointCells.GetCells(shape, VASS.CellValueType.Formula);
                dicof_shape_to_cxnpoint[shape] = cp;
            }

            return dicof_shape_to_cxnpoint;
        }

        public List<int> AddConnectionPoint(
            TargetShapes targets, 
            string fx,
            string fy,
            Models.ConnectionPointType type)
        {
            targets = targets.ResolveToShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return new List<int>(0);
            }

            int dirx = 0;
            int diry = 0;

            var indices = new List<int>(targets.Shapes.Count);

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(AddConnectionPoint)))
            {
                var cxnpointcells = new VA.Shapes.ConnectionPointCells();
                cxnpointcells.X = fx;
                cxnpointcells.Y = fy;
                cxnpointcells.DirX = dirx;
                cxnpointcells.DirY = diry;
                cxnpointcells.Type = (int)type;

                foreach (var shape in targets.Shapes)
                {
                    int index = VA.Shapes.ConnectionPointHelper.Add(shape, cxnpointcells);
                    indices.Add(index);
                }
            }

            return indices;
        }

        public void DeleteConnectionPointAtIndex(TargetShapes targetshapes, int index)
        {
            targetshapes = targetshapes.ResolveToShapes(this._client);

            if (targetshapes.Shapes.Count < 1)
            {
                return;
            }

            // restrict the operation to those shapes that actually have enough
            // connection points to qualify for deleting 
            var qualified_shapes = targetshapes.Shapes.Where(shape => VA.Shapes.ConnectionPointHelper.GetCount(shape) > index);

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DeleteConnectionPointAtIndex)))
            {
                foreach (var shape in qualified_shapes)
                {
                    VA.Shapes.ConnectionPointHelper.Delete(shape, index);
                }
            }
        }
    }
}
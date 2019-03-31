using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Shapes;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class ConnectionPointCommands : CommandSet
    {
        internal ConnectionPointCommands(Client client) :
            base(client)
        {

        }

        public IDictionary<IVisio.Shape, IList<ConnectionPointCells>> GetConnectionPoints(Models.TargetShapes targetshapes)
        {
            targetshapes = targetshapes.ResolveShapes(this._client);

            if (targetshapes.Shapes.Count<1)
            {
                return new Dictionary<IVisio.Shape, IList<ConnectionPointCells>>();
            }

            var dic = new Dictionary<IVisio.Shape, IList<ConnectionPointCells>>();
            foreach (var shape in targetshapes.Shapes)
            {
                var cp = VisioAutomation.Shapes.ConnectionPointCells.GetCells(shape, CellValueType.Formula);
                dic[shape] = cp;
            }

            return dic;
        }

        public List<int> AddConnectionPoint(
            Models.TargetShapes targets, 
            string fx,
            string fy,
            Models.ConnectionPointType type)
        {
            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return new List<int>(0);
            }

            int dirx = 0;
            int diry = 0;

            var indices = new List<int>(targets.Shapes.Count);

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(AddConnectionPoint)))
            {
                var cp = new ConnectionPointCells();
                cp.X = fx;
                cp.Y = fy;
                cp.DirX = dirx;
                cp.DirY = diry;
                cp.Type = (int)type;

                foreach (var shape in targets.Shapes)
                {
                    int index = ConnectionPointHelper.Add(shape, cp);
                    indices.Add(index);
                }
            }

            return indices;
        }
        
        public List<int> AddConnectionPoint(
            string fx,
            string fy,
            Models.ConnectionPointType type)
        {
            var targets = new Models.TargetShapes();
            return this.AddConnectionPoint(targets, fx, fy, type);
        }

        public void DeleteConnectionPointAtIndex(Models.TargetShapes targetshapes, int index)
        {
            targetshapes = targetshapes.ResolveShapes(this._client);

            if (targetshapes.Shapes.Count < 1)
            {
                return;
            }

            // restrict the operation to those shapes that actually have enough
            // connection points to qualify for deleting 
            var qualified_shapes = targetshapes.Shapes.Where(shape => ConnectionPointHelper.GetCount(shape) > index);

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DeleteConnectionPointAtIndex)))
            {
                foreach (var shape in qualified_shapes)
                {
                    ConnectionPointHelper.Delete(shape, index);
                }
            }
        }
    }
}
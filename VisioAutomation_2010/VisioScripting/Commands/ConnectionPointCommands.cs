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

        public IDictionary<IVisio.Shape, IList<ConnectionPointCells>> GetConnectionPoints(Models.TargetShapes targets)
        {
            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count<1)
            {
                return new Dictionary<IVisio.Shape, IList<ConnectionPointCells>>();
            }

            var dic = new Dictionary<IVisio.Shape, IList<ConnectionPointCells>>();
            foreach (var shape in targets.Shapes)
            {
                var cp = ConnectionPointCells.GetCells(shape, CellValueType.Formula);
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

            using (var undoscope = this._client.Application.NewUndoScope(nameof(AddConnectionPoint)))
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

        public void DeleteConnectionPointAtIndex(Models.TargetShapes targets, int index)
        {
            var shapes = targets.ResolveShapes(this._client);

            if (shapes.Shapes.Count < 1)
            {
                return;
            }

            // restrict the operation to those shapes that actually have enough
            // connection points to qualify for deleting 
            var qualified_shapes = shapes.Shapes.Where(shape => ConnectionPointHelper.GetCount(shape) > index);

            using (var undoscope = this._client.Application.NewUndoScope(nameof(DeleteConnectionPointAtIndex)))
            {
                foreach (var shape in qualified_shapes)
                {
                    ConnectionPointHelper.Delete(shape, index);
                }
            }
        }
    }
}
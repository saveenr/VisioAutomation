using System.Collections.Generic;
using System.Linq;
using VisioScripting.Models;
using VisioAutomation.Shapes;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class ConnectionPointCommands : CommandSet
    {
        internal ConnectionPointCommands(Client client) :
            base(client)
        {

        }

        public IDictionary<IVisio.Shape, IList<ConnectionPointCells>> Get(TargetShapes targets)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count<1)
            {
                return new Dictionary<IVisio.Shape, IList<ConnectionPointCells>>();
            }

            var dic = new Dictionary<IVisio.Shape, IList<ConnectionPointCells>>();
            foreach (var shape in targets.Shapes)
            {
                var cp = ConnectionPointCells.GetCells(shape);
                dic[shape] = cp;
            }

            return dic;
        }

        public List<int> Add( TargetShapes targets, 
            string fx,
            string fy,
            ConnectionPointType type)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return new List<int>(0);
            }

            int dirx = 0;
            int diry = 0;

            var indices = new List<int>(targets.Shapes.Count);

            using (var undoscope = this._client.Application.NewUndoScope("Add Connection Point"))
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


        public List<int> Add(
            string fx,
            string fy,
            ConnectionPointType type)
        {
            this._client.Application.AssertApplicationAvailable();

            var targets = new TargetShapes();
            return this.Add(targets, fx, fy, type);
        }

        public void Delete(TargetShapes targets, int index)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapes(this._client);

            if (shapes.Shapes.Count < 1)
            {
                return;
            }

            var target_shapes = shapes.Shapes.Where(shape => ConnectionPointHelper.GetCount(shape) > index);

            using (var undoscope = this._client.Application.NewUndoScope("Delete Connection Point"))
            {
                foreach (var shape in target_shapes)
                {
                    ConnectionPointHelper.Delete(shape, index);
                }
            }
        }
    }
}
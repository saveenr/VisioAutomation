using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Shapes.ConnectionPoints;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Commands
{
    public class ConnectionPointCommands : CommandSet
    {
        internal ConnectionPointCommands(Client client) :
            base(client)
        {

        }

        public IDictionary<IVisio.Shape, IList<ConnectionPointCells>> Get(IList<IVisio.Shape> target_shapes)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);

            if (shapes.Count<1)
            {
                return new Dictionary<IVisio.Shape, IList<ConnectionPointCells>>();
            }

            var dic = new Dictionary<IVisio.Shape, IList<ConnectionPointCells>>();
            foreach (var shape in shapes)
            {
                var cp = ConnectionPointCells.GetCells(shape);
                dic[shape] = cp;
            }

            return dic;
        }

        public IList<int> Add( IList<IVisio.Shape> target_shapes, 
            string fx,
            string fy,
            ConnectionPointType type)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return new List<int>(0);
            }

            int dirx = 0;
            int diry = 0;

            var indices = new List<int>(shapes.Count);

            var app = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Add Connection Point"))
            {
                var cp = new ConnectionPointCells();
                cp.X = fx;
                cp.Y = fy;
                cp.DirX = dirx;
                cp.DirY = diry;
                cp.Type = (int)type;

                foreach (var shape in shapes)
                {
                    int index = ConnectionPointHelper.Add(shape, cp);
                    indices.Add(index);
                }
            }

            return indices;
        }


        public IList<int> Add(
            string fx,
            string fy,
            ConnectionPointType type)
        {
            this._client.Application.AssertApplicationAvailable();

            return this.Add(null, fx, fy, type);
        }

        public void Delete(List<IVisio.Shape> target_shapes0, int index)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes0);
            if (shapes.Count < 1)
            {
                return;
            }

            var target_shapes = shapes.Where(shape => ConnectionPointHelper.GetCount(shape) > index);

            var application = this._client.Application.Get();
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
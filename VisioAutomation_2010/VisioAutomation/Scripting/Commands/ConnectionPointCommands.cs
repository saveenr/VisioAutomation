using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using CONS = VisioAutomation.Shapes.Connections;

namespace VisioAutomation.Scripting.Commands
{
    public class ConnectionPointCommands : CommandSet
    {
        internal ConnectionPointCommands(Client client) :
            base(client)
        {

        }

        public IDictionary<IVisio.Shape, IList<CONS.ConnectionPointCells>> Get(IList<IVisio.Shape> target_shapes)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);

            if (shapes.Count<1)
            {
                return new Dictionary<IVisio.Shape, IList<CONS.ConnectionPointCells>>();
            }

            var dic = new Dictionary<IVisio.Shape, IList<CONS.ConnectionPointCells>>();
            foreach (var shape in shapes)
            {
                var cp = CONS.ConnectionPointCells.GetCells(shape);
                dic[shape] = cp;
            }

            return dic;
        }

        public IList<int> Add( IList<IVisio.Shape> target_shapes, 
            string fx,
            string fy,
            CONS.ConnectionPointType type)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return new List<int>(0);
            }

            int dirx = 0;
            int diry = 0;

            var indices = new List<int>(shapes.Count);

            var app = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Add Connection Point"))
            {
                var cp = new CONS.ConnectionPointCells();
                cp.X = fx;
                cp.Y = fy;
                cp.DirX = dirx;
                cp.DirY = diry;
                cp.Type = (int)type;

                foreach (var shape in shapes)
                {
                    int index = CONS.ConnectionPointHelper.Add(shape, cp);
                    indices.Add(index);
                }
            }

            return indices;
        }


        public IList<int> Add(
            string fx,
            string fy,
            CONS.ConnectionPointType type)
        {
            this.Client.Application.AssertApplicationAvailable();

            return this.Add(null, fx, fy, type);
        }

        public void Delete(List<IVisio.Shape> target_shapes0, int index)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes0);
            if (shapes.Count < 1)
            {
                return;
            }

            var target_shapes = shapes.Where(shape => CONS.ConnectionPointHelper.GetCount(shape) > index);

            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Delete Connection Point"))
            {
                foreach (var shape in target_shapes)
                {
                    CONS.ConnectionPointHelper.Delete(shape, index);
                }
            }
        }
    }
}
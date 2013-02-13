using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class ConnectionPointCommands : CommandSet
    {
        public ConnectionPointCommands(Session session) :
            base(session)
        {

        }

        public IDictionary<IVisio.Shape, IList<VA.Connections.ConnectionPointCells>> Get(IList<IVisio.Shape> target_shapes)
        {
            var shapes = get_target_shapes(target_shapes);
            if (shapes.Count<1)
            {
                return new Dictionary<IVisio.Shape, IList<VA.Connections.ConnectionPointCells>>();
            }

            var dic = new Dictionary<IVisio.Shape, IList<VA.Connections.ConnectionPointCells>>();
            var application = this.Session.VisioApplication;
            foreach (var shape in shapes)
            {
                var cp = VA.Connections.ConnectionPointCells.GetCells(shape);
                dic[shape] = cp;
            }

            return dic;
        }

        public IDictionary<IVisio.Shape, IList<VA.Connections.ConnectionPointCells>> Get()
        {
            return this.Get(null);
        }

        public IList<int> Add( IList<IVisio.Shape> target_shapes, 
    string fx,
    string fy,
    VA.Connections.ConnectionPointType type)
        {
            var shapes = get_target_shapes(target_shapes);
            if (shapes.Count < 1)
            {
                return new List<int>(0);
            }


            int dirx = 0;
            int diry = 0;


            var indices = new List<int>(shapes.Count);

            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication, "Add Connection Point"))
            {
                var cp = new VA.Connections.ConnectionPointCells();
                cp.X = fx;
                cp.Y = fy;
                cp.DirX = dirx;
                cp.DirY = diry;
                cp.Type = (int)type;

                foreach (var shape in shapes)
                {

                    int index = VA.Connections.ConnectionPointHelper.Add(shape, cp);
                    indices.Add(index);
                }
            }

            return indices;
        }


        public IList<int> Add(
            string fx,
            string fy,
            VA.Connections.ConnectionPointType type)
        {
            return this.Add(null, fx, fy, type);
        }

        public void Delete(int index)
        {
            this.Delete(null,index);
        }

        public void Delete(List<IVisio.Shape> target_shapes0, int index)
        {
            var shapes = get_target_shapes(target_shapes0);
            if (shapes.Count < 1)
            {
                return;
            }

            var target_shapes = from shape in shapes
                                where VA.Connections.ConnectionPointHelper.GetCount(shape) > index
                                select shape;

            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication, "Delete Connection Point"))
            {
                foreach (var shape in target_shapes)
                {
                    VA.Connections.ConnectionPointHelper.Delete(shape, index);
                }
            }
        }
    }
}
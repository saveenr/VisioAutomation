using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class ControlCommands : CommandSet
    {
        public ControlCommands(Session session) :
            base(session)
        {

        }

        public IList<int> Add(VA.Controls.ControlCells ctrl)
        {
            return this.Add(null, ctrl);
        }

        public IList<int> Add(IList<IVisio.Shape> target_shapes, VA.Controls.ControlCells ctrl)
        {
            if (ctrl == null)
            {
                throw new System.ArgumentNullException("ctrl");
            }

            var shapes = get_target_shapes(target_shapes);
            if (shapes.Count < 1)
            {
                return new List<int>(0);
            }


            var control_indices = new List<int>();
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication,"Add Control"))
            {
                foreach (var shape in shapes)
                {
                    int ci = VA.Controls.ControlHelper.Add(shape, ctrl);
                    control_indices.Add(ci);
                }
            }

            return control_indices;
        }

        public void Delete(int n)
        {
            this.Delete(null,n);
        }

        public void Delete(IList<IVisio.Shape> target_shapes, int n)
        {
            var shapes = get_target_shapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication, "Delete Control"))
            {
                foreach (var shape in shapes)
                {
                    VA.Controls.ControlHelper.Delete(shape, n);
                }
            }
        }

        public Dictionary<IVisio.Shape, IList<VA.Controls.ControlCells>> Get()
        {
            return this.Get(null);
        }

        public Dictionary<IVisio.Shape, IList<VA.Controls.ControlCells>> Get(IList<IVisio.Shape> target_shapes)
        {
            var shapes = get_target_shapes(target_shapes);
            if (shapes.Count < 1)
            {
                return new Dictionary<IVisio.Shape, IList<VA.Controls.ControlCells>>(0);
            }

            var dic = new Dictionary<IVisio.Shape, IList<VA.Controls.ControlCells>>();
            foreach (var shape in shapes)
            {
                var controls = VA.Controls.ControlCells.GetCells(shape);
                dic[shape] = controls;
            }
            return dic;
        }
    }
}
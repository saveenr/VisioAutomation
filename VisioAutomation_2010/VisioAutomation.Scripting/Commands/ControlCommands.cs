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

        public IList<int> Add(IList<IVisio.Shape> target_shapes, VA.Controls.ControlCells ctrl)
        {
            this.CheckApplication();

            if (ctrl == null)
            {
                throw new System.ArgumentNullException("ctrl");
            }

            var shapes = GetTargetShapes(target_shapes);
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

        public void Delete(IList<IVisio.Shape> target_shapes, int n)
        {
            this.CheckApplication();

            var shapes = GetTargetShapes(target_shapes);
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

        public Dictionary<IVisio.Shape, IList<VA.Controls.ControlCells>> Get(IList<IVisio.Shape> target_shapes)
        {
            this.CheckApplication();
            
            var shapes = GetTargetShapes(target_shapes);
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
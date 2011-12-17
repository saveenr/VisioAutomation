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

        public IList<int> Add()
        {
            if (!this.Session.HasSelectedShapes())
            {
                return null;
            }

            var ctrl = new VA.Controls.ControlCells();
            var control_indices = Add(ctrl);

            return control_indices;
        }

        public IList<int> Add(VA.Controls.ControlCells ctrl)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return null;
            }

            if (ctrl == null)
            {
                throw new System.ArgumentNullException("ctrl");
            }

            var shapes = this.Session.Selection.EnumShapes().ToList();
            var control_indices = new List<int>();
            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                foreach (var shape in shapes)
                {
                    int ci = VA.Controls.ControlHelper.AddControl(shape, ctrl);
                    control_indices.Add(ci);
                }
            }

            return control_indices;
        }

        public void Delete(int n)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var shapes = this.Session.Selection.EnumShapes().ToList();

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                foreach (var shape in shapes)
                {
                    VA.Controls.ControlHelper.DeleteControl(shape, n);
                }
            }
        }

        public Dictionary<IVisio.Shape, IList<VA.Controls.ControlCells>> Get()
        {
            if (!this.Session.HasSelectedShapes())
            {
                return new Dictionary<IVisio.Shape, IList<VA.Controls.ControlCells>>(0);
            }

            var shapes = this.Session.Selection.EnumShapes().ToList();

            var dic = new Dictionary<IVisio.Shape, IList<VA.Controls.ControlCells>>();
            foreach (var shape in shapes)
            {
                var controls = VA.Controls.ControlHelper.GetControls(shape);
                dic[shape] = controls;
            }
            return dic;
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Controls;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class ControlCommands : SessionCommands
    {
        public ControlCommands(Session session) :
            base(session)
        {

        }

        public IList<int> AddControl()
        {
            if (!this.Session.HasSelectedShapes())
            {
                return null;
            }

            var ctrl = new VA.Controls.ControlCells();
            var control_indices = AddControl(ctrl);

            return control_indices;
        }

        public IList<int> AddControl(VA.Controls.ControlCells ctrl)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return null;
            }

            if (ctrl == null)
            {
                throw new ArgumentNullException("ctrl");
            }

            var shapes = this.Session.Selection.EnumSelectedShapes().ToList();
            var control_indices = new List<int>();
            var application = this.Session.Application;
            using (var undoscope = application.CreateUndoScope())
            {
                foreach (var shape in shapes)
                {
                    int ci = ControlHelper.AddControl(shape, ctrl);
                    control_indices.Add(ci);
                }
            }

            return control_indices;
        }

        public void DeleteControl(int n)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var shapes = this.Session.Selection.EnumSelectedShapes().ToList();

            var application = this.Session.Application;
            using (var undoscope = application.CreateUndoScope())
            {
                foreach (var shape in shapes)
                {
                    ControlHelper.DeleteControl(shape, n);
                }
            }
        }

        public IDictionary<IVisio.Shape, IList<VA.Controls.ControlCells>> GetControls()
        {
            if (!this.Session.HasSelectedShapes())
            {
                return new Dictionary<IVisio.Shape, IList<VA.Controls.ControlCells>>(0);
            }

            var shapes = this.Session.Selection.EnumSelectedShapes().ToList();

            var dic = new Dictionary<IVisio.Shape, IList<VA.Controls.ControlCells>>();
            foreach (var shape in shapes)
            {
                var controls = ControlHelper.GetControls(shape);
                dic[shape] = controls;
            }
            return dic;
        }

        public IList<IVisio.Shape> GetSubSelectedShapes()
        {
            //http://www.visguy.com/2008/05/17/detect-sub-selected-shapes-programmatically/
            var shapes = new List<IVisio.Shape>(0);
            var sel = this.Session.Selection.GetSelection();
            var original_itermode = sel.IterationMode;
            
            // normal selection
            sel.IterationMode = ((short)IVisio.VisSelectMode.visSelModeSkipSub) + ((short)IVisio.VisSelectMode.visSelModeSkipSuper);

            if (sel.Count > 0)
            {
                shapes.AddRange(sel.AsEnumerable());
            }

            // sub selection
            sel.IterationMode = ((short)IVisio.VisSelectMode.visSelModeOnlySub) + ((short)IVisio.VisSelectMode.visSelModeSkipSuper);
            if (sel.Count > 0)
            {
                shapes.AddRange(sel.AsEnumerable());
            }

            sel.IterationMode = original_itermode;
            return shapes;
        }
    }
}
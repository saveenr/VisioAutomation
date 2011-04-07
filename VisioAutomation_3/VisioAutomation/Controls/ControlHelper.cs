using System;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;

namespace VisioAutomation.Controls
{
    public static class ControlHelper
    {
        internal readonly static VA.Controls.ControlQuery query = new VA.Controls.ControlQuery();

        public static int AddControl(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            var ctrl = new ControlCells();

            return AddControl(shape, ctrl);
        }

        public static int AddControl(
            IVisio.Shape shape,
            ControlCells ctrl)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            short row = shape.AddRow((short)IVisio.VisSectionIndices.visSectionControls,
                                     (short)IVisio.VisRowIndices.visRowLast,
                                     (short)IVisio.VisRowTags.visTagDefault);

            SetControl(shape, row, ctrl);

            return row;
        }

        public static int SetControl(
            IVisio.Shape shape,
            short row,
            ControlCells ctrl)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }


            if (!ctrl.XDynamics.Formula.HasValue)
            {
                ctrl.XDynamics = String.Format("Controls.Row_{0}", row + 1);
            }

            if (!ctrl.YDynamics.Formula.HasValue)
            {
                ctrl.YDynamics = String.Format("Controls.Row_{0}.Y", row + 1);
            }

            var update = new VA.ShapeSheet.Update.SRCUpdate();
            ctrl.Apply(update, row);
            update.Execute(shape);

            return row;
        }

        public static void DeleteControl(IVisio.Shape shape, int index)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            if (index < 0)
            {
                throw new ArgumentOutOfRangeException("index");
            }

            var row = (IVisio.VisRowIndices)index;
            shape.DeleteRow(query.Section, (short)row);
        }

        public static int GetControlsCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            return shape.RowCount[query.Section];
        }

        public static IList<ControlCells> GetControls(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            var r = query.GetFormulasAndResults<double>(shape);
            var formulas = r.Formulas;
            var results = r.Results;

            var controls = new List<ControlCells>(formulas.Rows.Count);

            for (int row = 0; row < results.Rows.Count; row++)
            {
                var control = new ControlCells();

                control.X = r.GetItem(row, query.X);
                control.Y = r.GetItem(row, query.Y);
                control.XDynamics = r.GetItem(row, query.XDyn, v => (int)v);
                control.YDynamics = r.GetItem(row, query.YDyn, v => (int)v);
                control.XBehavior = r.GetItem(row, query.XCon, v => (int)v);
                control.YBehavior = r.GetItem(row, query.YCon, v => (int)v);
                control.CanGlue = r.GetItem(row, query.Glue, v => (int)v);
                control.Tip = r.GetItem(row, query.Tip, v => (int)v);

                controls.Add(control);
            }

            return controls;
        }
    }
}
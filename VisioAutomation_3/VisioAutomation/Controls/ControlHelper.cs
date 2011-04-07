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
            return ControlCells.GetCells(shape);
        }

        public static IList<List<ControlCells>> GetControls(IVisio.Page page, IList<int> shapeids)
        {
            return ControlCells.GetCells(page,shapeids);
        }
    }
}
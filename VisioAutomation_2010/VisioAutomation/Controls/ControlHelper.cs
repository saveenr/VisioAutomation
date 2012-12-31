using System;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;

namespace VisioAutomation.Controls
{
    public static class ControlHelper
    {
        public static int Add(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            var ctrl = new ControlCells();

            return Add(shape, ctrl);
        }

        public static int Add(
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

            Set(shape, row, ctrl);

            return row;
        }

        public static int Set(
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

            var update = new VA.ShapeSheet.Update();
            update.SetFormulasForRow(ctrl, row);
            update.Execute(shape);

            return row;
        }

        public static void Delete(IVisio.Shape shape, int index)
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
            shape.DeleteRow( (short) IVisio.VisSectionIndices.visSectionControls, (short)row);
        }

        public static int GetCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            return shape.RowCount[(short)IVisio.VisSectionIndices.visSectionControls];
        }

        public static IList<ControlCells> Get(IVisio.Shape shape)
        {
            return ControlCells.GetCells(shape);
        }

        public static IList<List<ControlCells>> Get(IVisio.Page page, IList<int> shapeids)
        {
            return ControlCells.GetCells(page,shapeids);
        }
    }
}
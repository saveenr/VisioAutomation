using System;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public static class ControlHelper
    {
        public static int Add(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            var ctrl = new ControlCells();

            return ControlHelper.Add(shape, ctrl);
        }

        public static int Add(IVisio.Shape shape, ControlCells ctrl)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            short row = shape.AddRow((short)IVisio.VisSectionIndices.visSectionControls,
                                     (short)IVisio.VisRowIndices.visRowLast,
                                     (short)IVisio.VisRowTags.visTagDefault);

            ControlHelper.Set(shape, row, ctrl);

            return row;
        }

        public static int Set( IVisio.Shape shape, short row, ControlCells ctrl)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }


            if (ctrl.XDynamics.Value==null)
            {
                ctrl.XDynamics = string.Format("Controls.Row_{0}", row + 1);
            }

            if (ctrl.YDynamics.Value == null)
            {
                ctrl.YDynamics = string.Format("Controls.Row_{0}.Y", row + 1);
            }

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetFormulas(ctrl, row);

            writer.Commit(shape);

            return row;
        }

        public static void Delete(IVisio.Shape shape, int index)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            if (index < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            var row = (IVisio.VisRowIndices)index;
            shape.DeleteRow( (short) IVisio.VisSectionIndices.visSectionControls, (short)row);
        }

        public static int GetCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            return shape.RowCount[(short)IVisio.VisSectionIndices.visSectionControls];
        }


    }
}
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public static class ControlHelper
    {
        public static int Add(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            var ctrl = new ControlCells();

            return Add(shape, ctrl);
        }

        public static int Add(IVisio.Shape shape, ControlCells ctrl)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            short row = shape.AddRow((short)IVisio.VisSectionIndices.visSectionControls,
                                     (short)IVisio.VisRowIndices.visRowLast,
                                     (short)IVisio.VisRowTags.visTagDefault);

            Set(shape, row, ctrl);

            return row;
        }

        public static int Set( IVisio.Shape shape, short row, ControlCells ctrl)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }


            if (ctrl.XDynamics.Value==null)
            {
                ctrl.XDynamics = string.Format("Controls.Row_{0}", row + 1);
            }

            if (ctrl.YDynamics.Value == null)
            {
                ctrl.YDynamics = string.Format("Controls.Row_{0}.Y", row + 1);
            }

            var writer = new ShapeSheet.Writers.SrcWriter();
            writer.SetValues(ctrl, row);

            writer.Commit(shape, Core.CellValueType.Formula);

            return row;
        }

        public static void Delete(IVisio.Shape shape, int index)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            if (index < 0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(index));
            }

            var row = (IVisio.VisRowIndices)index;
            shape.DeleteRow( (short) IVisio.VisSectionIndices.visSectionControls, (short)row);
        }

        public static int GetCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            return shape.RowCount[(short)IVisio.VisSectionIndices.visSectionControls];
        }

    }
}
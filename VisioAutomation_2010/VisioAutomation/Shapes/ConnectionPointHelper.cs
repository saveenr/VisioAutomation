using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VASS=VisioAutomation.ShapeSheet;

namespace VisioAutomation.Shapes
{
    public static class ConnectionPointHelper
    {
        public static int Add(
            IVisio.Shape shape,
            ConnectionPointCells connection_point_cells)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            if (connection_point_cells.X.Value==null)
            {
                string msg = "Must provide an X Formula";
                throw new System.ArgumentException(msg, nameof(connection_point_cells));
            }

            if (connection_point_cells.Y.Value==null)
            {
                string msg = "Must provide an Y Formula";
                throw new System.ArgumentException(msg, nameof(connection_point_cells));
            }

            var n = shape.AddRow((short)IVisio.VisSectionIndices.visSectionConnectionPts,
                                 (short)IVisio.VisRowIndices.visRowLast,
                                 (short)IVisio.VisRowTags.visTagCnnctPt);

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(connection_point_cells, n);

            writer.CommitFormulas(shape);

            return n;
        }

        public static int Set(
            IVisio.Shape shape,
            short row,
            ConnectionPointCells cpcells)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(cpcells, row);

            writer.CommitFormulas(shape);

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
            shape.DeleteRow( (short) IVisio.VisSectionIndices.visSectionConnectionPts, (short)row);
        }

        public static int Delete(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            int n = ConnectionPointHelper.GetCount(shape);
            for (int i = n - 1; i >= 0; i--)
            {
                ConnectionPointHelper.Delete(shape, i);
            }

            return n;
        }


        public static int GetCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            return shape.RowCount[(short)IVisio.VisSectionIndices.visSectionConnectionPts];
        }

    }
}
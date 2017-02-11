using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.ConnectionPoints
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

            if (!connection_point_cells.X.Formula.HasValue)
            {
                string msg = "Must provide an X Formula";
                throw new System.ArgumentException(msg, nameof(connection_point_cells));
            }

            if (!connection_point_cells.Y.Formula.HasValue)
            {
                string msg = "Must provide an Y Formula";
                throw new System.ArgumentException(msg, nameof(connection_point_cells));
            }

            var n = shape.AddRow((short)IVisio.VisSectionIndices.visSectionConnectionPts,
                                 (short)IVisio.VisRowIndices.visRowLast,
                                 (short)IVisio.VisRowTags.visTagCnnctPt);

            var writer = new ShapeSheet.Writer.ShapeSheetWriter();
            connection_point_cells.SetFormulas(writer,n);

            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(shape);
            writer.Commit(surface);

            return n;
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

        public static int GetCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            return shape.RowCount[ (short) IVisio.VisSectionIndices.visSectionConnectionPts];
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
    }
}
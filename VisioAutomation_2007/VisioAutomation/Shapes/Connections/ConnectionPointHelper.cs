using VA=VisioAutomation;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Shapes.Connections
{
    public static class ConnectionPointHelper
    {
        public static int Add(
            IVisio.Shape shape,
            ConnectionPointCells cp)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            if (!cp.X.Formula.HasValue)
            {
                throw new System.ArgumentException("Must provide an X Formula");
            }

            if (!cp.Y.Formula.HasValue)
            {
                throw new System.ArgumentException("Must provide an Y Formula");
            }

            var n = shape.AddRow((short)IVisio.VisSectionIndices.visSectionConnectionPts,
                                 (short)IVisio.VisRowIndices.visRowLast,
                                 (short)IVisio.VisRowTags.visTagCnnctPt);

            var update = new VA.ShapeSheet.Update();
            update.SetFormulas(cp,n);
            update.Execute(shape);

            return n;
        }

        public static void Delete(IVisio.Shape shape, int index)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            if (index < 0)
            {
                throw new System.ArgumentOutOfRangeException("index");
            }

            var row = (IVisio.VisRowIndices)index;
            shape.DeleteRow( (short) IVisio.VisSectionIndices.visSectionConnectionPts, (short)row);
        }

        public static int GetCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            return shape.RowCount[ (short) IVisio.VisSectionIndices.visSectionConnectionPts];
        }

        public static int Delete(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            int n = GetCount(shape);
            for (int i = n - 1; i >= 0; i--)
            {
                Delete(shape, i);
            }

            return n;
        }
    }
}
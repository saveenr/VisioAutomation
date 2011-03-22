using Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Connections
{
    public enum ConnectionPointType
    {
        Inward = IVisio.VisCellVals.visCnnctTypeInward,
        Outward = IVisio.VisCellVals.visCnnctTypeOutward,
        InwardOutward = IVisio.VisCellVals.visCnnctTypeInwardOutward
    }

    public class ConnectionPointCells
    {
        public VA.ShapeSheet.CellData<double> X { get; set; }
        public VA.ShapeSheet.CellData<double> Y { get; set; }
        public VA.ShapeSheet.CellData<int> DirX { get; set; }
        public VA.ShapeSheet.CellData<int> DirY { get; set; }
        public VA.ShapeSheet.CellData<int> Type { get; set; }


        internal readonly static VA.Connections.ConnectionPointQuery query = new VA.Connections.ConnectionPointQuery();

        public static IList<ConnectionPointCells> GetConnectionPoints(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            var qds = query.GetFormulasAndResults<double>(shape);

            var connectionpoints = new List<ConnectionPointCells>(qds.RowCount);
            for (int row = 0; row < qds.RowCount; row++)
            {
                var connectionpoint = new ConnectionPointCells();
                connectionpoint.X = qds.GetItem(row, query.X);
                connectionpoint.Y = qds.GetItem(row, query.Y);
                connectionpoint.DirX = qds.GetItem(row, query.DirX, v => (int)v);
                connectionpoint.DirY = qds.GetItem(row, query.DirY, v => (int)v);
                connectionpoint.Type = qds.GetItem(row, query.Type, v => (int)v);
                connectionpoints.Add(connectionpoint);
            }

            return connectionpoints;
        }

        public void Apply(VA.ShapeSheet.Update.SRCUpdate update, short n)
        {
            var cp = this;
            var src_x = ConnectionPointCells.query.GetCellSRCForRow(ConnectionPointCells.query.X, n);
            var src_y = ConnectionPointCells.query.GetCellSRCForRow(ConnectionPointCells.query.Y, n);
            var src_dirx = ConnectionPointCells.query.GetCellSRCForRow(ConnectionPointCells.query.DirX, n);
            var src_diry = ConnectionPointCells.query.GetCellSRCForRow(ConnectionPointCells.query.DirY, n);
            var src_type = ConnectionPointCells.query.GetCellSRCForRow(ConnectionPointCells.query.Type, n);

            update.SetFormula(src_x, cp.X.Formula);
            update.SetFormula(src_y, cp.Y.Formula);
            update.SetFormulaIgnoreNull(src_dirx, cp.DirX.Formula);
            update.SetFormulaIgnoreNull(src_diry, cp.DirY.Formula);
            update.SetFormulaIgnoreNull(src_type, cp.Type.Formula);
        }

    }

    public static class ConnectionPointHelper
    {
        public static int AddConnectionPoint(
            IVisio.Shape shape,
            Connections.ConnectionPointCells cp)
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

            var n = shape.AddRow((short)VisSectionIndices.visSectionConnectionPts,
                                 (short)VisRowIndices.visRowLast,
                                 (short)VisRowTags.visTagCnnctPt);

            var update = new VA.ShapeSheet.Update.SRCUpdate();
            cp.Apply(update,n);
            update.Execute(shape);

            return n;
        }


        public static void DeleteConnectionPoint(IVisio.Shape shape, int index)
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
            shape.DeleteRow(ConnectionPointCells.query.Section, (short)row);
        }

        public static int GetConnectionPointCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            return shape.RowCount[ConnectionPointCells.query.Section];
        }

        public static int DeleteAllConnectionPoints(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            int n = GetConnectionPointCount(shape);
            for (int i = n - 1; i >= 0; i--)
            {
                DeleteConnectionPoint(shape, i);
            }

            return n;
        }
    }


}
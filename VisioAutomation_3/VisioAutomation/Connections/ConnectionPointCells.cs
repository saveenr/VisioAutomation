using VA=VisioAutomation;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Connections
{
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
}
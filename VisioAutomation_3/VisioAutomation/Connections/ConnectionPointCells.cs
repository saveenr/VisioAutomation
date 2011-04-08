using VA=VisioAutomation;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Connections
{
    public class ConnectionPointCells : VA.ShapeSheet.CellSectionDataGroup
    {
        public VA.ShapeSheet.CellData<double> X { get; set; }
        public VA.ShapeSheet.CellData<double> Y { get; set; }
        public VA.ShapeSheet.CellData<int> DirX { get; set; }
        public VA.ShapeSheet.CellData<int> DirY { get; set; }
        public VA.ShapeSheet.CellData<int> Type { get; set; }


        internal readonly static VA.Connections.ConnectionPointQuery query = new VA.Connections.ConnectionPointQuery();

        [System.Obsolete]
        public static IList<ConnectionPointCells> GetConnectionPoints(IVisio.Shape shape)
        {
            return ConnectionPointCells.GetCells(shape);
        }

        protected override void _Apply(VA.ShapeSheet.CellSectionDataGroup.ApplyFormula func, short row)
        {
            func(VA.ShapeSheet.SRCConstants.Connections_X.ForRow(row), this.X.Formula);
            func(VA.ShapeSheet.SRCConstants.Connections_Y.ForRow(row), this.Y.Formula);
            func(VA.ShapeSheet.SRCConstants.Connections_DirX.ForRow(row), this.DirX.Formula);
            func(VA.ShapeSheet.SRCConstants.Connections_DirY.ForRow(row), this.DirY.Formula);
            func(VA.ShapeSheet.SRCConstants.Connections_Type.ForRow(row), this.Type.Formula);
        }

        private static ConnectionPointCells get_cells_from_row(ConnectionPointQuery query, VA.ShapeSheet.Query.QueryDataSet<double> qds, int row)
        {
            var cells = new ConnectionPointCells();
            cells.X = qds.GetItem(row, query.X);
            cells.Y = qds.GetItem(row, query.Y);
            cells.DirX = qds.GetItem(row, query.DirX, v => (int)v);
            cells.DirY = qds.GetItem(row, query.DirY, v => (int)v);
            cells.Type = qds.GetItem(row, query.Type, v => (int)v);

            return cells;
        }

        public static IList<List<ConnectionPointCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new ConnectionPointQuery();
            return VA.ShapeSheet.CellSectionDataGroup._GetCells(page, shapeids, query, get_cells_from_row);
        }

        public static IList<ConnectionPointCells> GetCells(IVisio.Shape shape)
        {
            var query = new ConnectionPointQuery();
            return VA.ShapeSheet.CellSectionDataGroup._GetCells(shape, query, get_cells_from_row);
        }
    }
}
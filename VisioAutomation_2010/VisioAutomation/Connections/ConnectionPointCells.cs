using VA=VisioAutomation;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;
using VisioAutomation.Extensions;

namespace VisioAutomation.Connections
{
    public class ConnectionPointCells : VA.ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public VA.ShapeSheet.CellData<double> X { get; set; }
        public VA.ShapeSheet.CellData<double> Y { get; set; }
        public VA.ShapeSheet.CellData<int> DirX { get; set; }
        public VA.ShapeSheet.CellData<int> DirY { get; set; }
        public VA.ShapeSheet.CellData<int> Type { get; set; }

        public override void ApplyFormulasForRow(ApplyFormula func, short row)
        {
            func(VA.ShapeSheet.SRCConstants.Connections_X.ForRow(row), this.X.Formula);
            func(VA.ShapeSheet.SRCConstants.Connections_Y.ForRow(row), this.Y.Formula);
            func(VA.ShapeSheet.SRCConstants.Connections_DirX.ForRow(row), this.DirX.Formula);
            func(VA.ShapeSheet.SRCConstants.Connections_DirY.ForRow(row), this.DirY.Formula);
            func(VA.ShapeSheet.SRCConstants.Connections_Type.ForRow(row), this.Type.Formula);
        }

        private static ConnectionPointCells get_cells_from_row(ConnectionPointQuery query, VA.ShapeSheet.Data.Table<VA.ShapeSheet.CellData<double>> table, int row)
        {
            var cells = new ConnectionPointCells();
            cells.X = table[row,query.X];
            cells.Y = table[row,query.Y];
            cells.DirX = table[row,query.DirX].ToInt();
            cells.DirY = table[row,query.DirY].ToInt();
            cells.Type = table[row,query.Type].ToInt();

            return cells;
        }

        public static IList<List<ConnectionPointCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroupMultiRow.CellsFromRowsGrouped(page, shapeids, query, get_cells_from_row);
        }

        public static IList<ConnectionPointCells> GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroupMultiRow.CellsFromRows(shape, query, get_cells_from_row);
        }

        private static ConnectionPointQuery m_query;
        private static ConnectionPointQuery get_query()
        {
            if (m_query == null)
            {
                m_query = new ConnectionPointQuery();
            }
            return m_query;
        }

        class ConnectionPointQuery : VA.ShapeSheet.Query.SectionQuery
        {
            public VA.ShapeSheet.Query.QueryColumn DirX { get; set; }
            public VA.ShapeSheet.Query.QueryColumn DirY { get; set; }
            public VA.ShapeSheet.Query.QueryColumn Type { get; set; }
            public VA.ShapeSheet.Query.QueryColumn X { get; set; }
            public VA.ShapeSheet.Query.QueryColumn Y { get; set; }

            public ConnectionPointQuery() :
                base(IVisio.VisSectionIndices.visSectionConnectionPts)
            {
                DirX = this.AddColumn(VA.ShapeSheet.SRCConstants.Connections_DirX, "DirX");
                DirY = this.AddColumn(VA.ShapeSheet.SRCConstants.Connections_DirY, "DirY");
                Type = this.AddColumn(VA.ShapeSheet.SRCConstants.Connections_Type, "Type");
                X = this.AddColumn(VA.ShapeSheet.SRCConstants.Connections_X, "X");
                Y = this.AddColumn(VA.ShapeSheet.SRCConstants.Connections_Y, "Y");
            }
        }
    }
}
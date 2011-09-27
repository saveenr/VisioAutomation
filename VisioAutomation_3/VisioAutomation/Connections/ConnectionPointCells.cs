using VA=VisioAutomation;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;
using VisioAutomation.Extensions;

namespace VisioAutomation.Connections
{
    public class ConnectionPointCells : VA.ShapeSheet.CellSectionDataGroup
    {
        public VA.ShapeSheet.CellData<double> X { get; set; }
        public VA.ShapeSheet.CellData<double> Y { get; set; }
        public VA.ShapeSheet.CellData<int> DirX { get; set; }
        public VA.ShapeSheet.CellData<int> DirY { get; set; }
        public VA.ShapeSheet.CellData<int> Type { get; set; }

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

        private static ConnectionPointCells get_cells_from_row2(ConnectionPointQuery query, VA.ShapeSheet.Data.QueryDataRow<double> row)
        {
            var cells = new ConnectionPointCells();
            cells.X = row[query.X];
            cells.Y = row[query.Y];
            cells.DirX = row[query.DirX].ToInt();
            cells.DirY = row[query.DirY].ToInt();
            cells.Type = row[query.Type].ToInt();

            return cells;
        }

        internal static IList<List<ConnectionPointCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new ConnectionPointQuery();
            return VA.ShapeSheet.CellSectionDataGroup._GetObjectsFromRowsGrouped(page, shapeids, query, get_cells_from_row2);
        }

        internal static IList<ConnectionPointCells> GetCells(IVisio.Shape shape)
        {
            var query = new ConnectionPointQuery();
            return VA.ShapeSheet.CellSectionDataGroup._GetObjectsFromRows(shape, query, get_cells_from_row2);
        }

        class ConnectionPointQuery : VA.ShapeSheet.Query.SectionQuery
        {
            public VA.ShapeSheet.Query.SectionQueryColumn DirX { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn DirY { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn Type { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn X { get; set; }
            public VA.ShapeSheet.Query.SectionQueryColumn Y { get; set; }

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
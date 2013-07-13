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
        
        public static IList<List<ConnectionPointCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return _GetCells(page, shapeids, query, query.GetCells);
        }

        public static IList<ConnectionPointCells> GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return _GetCells(shape, query, query.GetCells);
        }

        private static ConnectionPointQuery m_query;

        private static ConnectionPointQuery get_query()
        {
            m_query =  m_query ?? new ConnectionPointQuery();
            return m_query;
        }

        class ConnectionPointQuery : VA.ShapeSheet.Query.QueryEx
        {
            public int DirX { get; set; }
            public int DirY { get; set; }
            public int Type { get; set; }
            public int X { get; set; }
            public int Y { get; set; }

            public ConnectionPointQuery()
            {
                var sec = this.AddSection(IVisio.VisSectionIndices.visSectionConnectionPts);
                DirX = sec.AddColumn(VA.ShapeSheet.SRCConstants.Connections_DirX, "DirX");
                DirY = sec.AddColumn(VA.ShapeSheet.SRCConstants.Connections_DirY, "DirY");
                Type = sec.AddColumn(VA.ShapeSheet.SRCConstants.Connections_Type, "Type");
                X = sec.AddColumn(VA.ShapeSheet.SRCConstants.Connections_X, "X");
                Y = sec.AddColumn(VA.ShapeSheet.SRCConstants.Connections_Y, "Y");
            }

            public ConnectionPointCells GetCells(VA.ShapeSheet.CellData<double>[] row)
            {
                var cells = new ConnectionPointCells();
                cells.X = row[this.X];
                cells.Y = row[this.Y];
                cells.DirX = row[this.DirX].ToInt();
                cells.DirY = row[this.DirY].ToInt();
                cells.Type = row[this.Type].ToInt();

                return cells;
            }
        }
    }
}
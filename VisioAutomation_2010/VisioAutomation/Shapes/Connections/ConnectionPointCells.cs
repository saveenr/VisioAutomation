using VA=VisioAutomation;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes.Connections
{
    public class ConnectionPointCells : VA.ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public VA.ShapeSheet.CellData<double> X { get; set; }
        public VA.ShapeSheet.CellData<double> Y { get; set; }
        public VA.ShapeSheet.CellData<int> DirX { get; set; }
        public VA.ShapeSheet.CellData<int> DirY { get; set; }
        public VA.ShapeSheet.CellData<int> Type { get; set; }

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return newpair(VA.ShapeSheet.SRCConstants.Connections_X, this.X.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Connections_Y, this.Y.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Connections_DirX, this.DirX.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Connections_DirY, this.DirY.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Connections_Type, this.Type.Formula);
            }
        }

        public static IList<List<ConnectionPointCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return _GetCells<ConnectionPointCells,double>(page, shapeids, query, query.GetCells);
        }

        public static IList<ConnectionPointCells> GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return _GetCells<ConnectionPointCells,double>(shape, query, query.GetCells);
        }

        private static ConnectionPointCellQuery _mCellQuery;

        private static ConnectionPointCellQuery get_query()
        {
            _mCellQuery =  _mCellQuery ?? new ConnectionPointCellQuery();
            return _mCellQuery;
        }

        class ConnectionPointCellQuery : VA.ShapeSheet.Query.CellQuery
        {
            public CellColumn DirX { get; set; }
            public CellColumn DirY { get; set; }
            public CellColumn Type { get; set; }
            public CellColumn X { get; set; }
            public CellColumn Y { get; set; }
            
            public ConnectionPointCellQuery()
            {
                var sec = this.AddSection(IVisio.VisSectionIndices.visSectionConnectionPts);
                DirX = sec.AddCell(VA.ShapeSheet.SRCConstants.Connections_DirX, "DirX");
                DirY = sec.AddCell(VA.ShapeSheet.SRCConstants.Connections_DirY, "DirY");
                Type = sec.AddCell(VA.ShapeSheet.SRCConstants.Connections_Type, "Type");
                X = sec.AddCell(VA.ShapeSheet.SRCConstants.Connections_X, "X");
                Y = sec.AddCell(VA.ShapeSheet.SRCConstants.Connections_Y, "Y");
            }

            public ConnectionPointCells GetCells(IList<VA.ShapeSheet.CellData<double>> row)
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
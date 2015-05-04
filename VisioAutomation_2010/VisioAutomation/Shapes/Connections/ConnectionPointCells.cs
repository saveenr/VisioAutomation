using VA=VisioAutomation;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes.Connections
{
    public class ConnectionPointCells : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public ShapeSheet.CellData<double> X { get; set; }
        public ShapeSheet.CellData<double> Y { get; set; }
        public ShapeSheet.CellData<int> DirX { get; set; }
        public ShapeSheet.CellData<int> DirY { get; set; }
        public ShapeSheet.CellData<int> Type { get; set; }

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SRCConstants.Connections_X, this.X.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Connections_Y, this.Y.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Connections_DirX, this.DirX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Connections_DirY, this.DirY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Connections_Type, this.Type.Formula);
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

        class ConnectionPointCellQuery : CellQuery
        {
            public CellColumn DirX { get; set; }
            public CellColumn DirY { get; set; }
            public CellColumn Type { get; set; }
            public CellColumn X { get; set; }
            public CellColumn Y { get; set; }
            
            public ConnectionPointCellQuery()
            {
                var sec = this.AddSection(IVisio.VisSectionIndices.visSectionConnectionPts);
                this.DirX = sec.AddCell(ShapeSheet.SRCConstants.Connections_DirX,"Connections_DirX");
                this.DirY = sec.AddCell(ShapeSheet.SRCConstants.Connections_DirY,"Connections_DirY");
                this.Type = sec.AddCell(ShapeSheet.SRCConstants.Connections_Type,"Connections_Type");
                this.X = sec.AddCell(ShapeSheet.SRCConstants.Connections_X,"Connections_X");
                this.Y = sec.AddCell(ShapeSheet.SRCConstants.Connections_Y,"Connections_Y");
            }

            public ConnectionPointCells GetCells(IList<ShapeSheet.CellData<double>> row)
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
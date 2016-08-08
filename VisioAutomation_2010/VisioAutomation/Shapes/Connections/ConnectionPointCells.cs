using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.Connections
{
    public class ConnectionPointCells : ShapeSheetQuery.CellGroups.CellGroupMultiRow
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
            var query = ConnectionPointCells.lazy_query.Value;
            return ShapeSheetQuery.CellGroups.CellGroupMultiRow._GetCells<ConnectionPointCells, double>(page, shapeids, query, query.GetCells);
        }

        public static IList<ConnectionPointCells> GetCells(IVisio.Shape shape)
        {
            var query = ConnectionPointCells.lazy_query.Value;
            return ShapeSheetQuery.CellGroups.CellGroupMultiRow._GetCells<ConnectionPointCells, double>(shape, query, query.GetCells);
        }

        private static System.Lazy<ShapeSheetQuery.Common.ConnectionPointCellsQuery> lazy_query = new System.Lazy<ShapeSheetQuery.Common.ConnectionPointCellsQuery>();


    }
}
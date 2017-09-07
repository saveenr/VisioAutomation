using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class ConnectionPointCells : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral X { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Y { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral DirX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral DirY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Type { get; set; }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.ConnectionPointX, this.X.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ConnectionPointY, this.Y.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ConnectionPointDirX, this.DirX.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ConnectionPointDirY, this.DirY.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ConnectionPointType, this.Type.Value);
            }
        }

        public static List<List<ConnectionPointCells>> GetCells(IVisio.Page page, IList<int> shapeids, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = ConnectionPointCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids, cvt);
        }

        public static List<ConnectionPointCells> GetCells(IVisio.Shape shape, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = ConnectionPointCells.lazy_query.Value;
            return query.GetCellGroups(shape, cvt);
        }

        private static readonly System.Lazy<ConnectionPointCellsReader> lazy_query = new System.Lazy<ConnectionPointCellsReader>();
    }
}
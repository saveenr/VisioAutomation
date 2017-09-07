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

        public static List<List<ConnectionPointCells>> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var query = ConnectionPointCells.lazy_query.Value;
            return query.GetFormulas(page, shapeids);
        }

        public static List<List<ConnectionPointCells>> GetResults(IVisio.Page page, IList<int> shapeids)
        {
            var query = ConnectionPointCells.lazy_query.Value;
            return query.GetResults(page, shapeids);
        }

        public static List<ConnectionPointCells> GetFormulas(IVisio.Shape shape)
        {
            var query = ConnectionPointCells.lazy_query.Value;
            return query.GetFormulas(shape);
        }

        public static List<ConnectionPointCells> GetResults(IVisio.Shape shape)
        {
            var query = ConnectionPointCells.lazy_query.Value;
            return query.GetResults(shape);
        }

        private static readonly System.Lazy<ConnectionPointCellsReader> lazy_query = new System.Lazy<ConnectionPointCellsReader>();
    }
}
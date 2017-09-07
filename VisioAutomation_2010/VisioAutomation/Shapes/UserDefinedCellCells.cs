using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class UserDefinedCellCells : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral Value { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Prompt { get; set; }

        public UserDefinedCellCells()
        {
        }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.UserDefCellValue, this.Value.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.UserDefCellPrompt, this.Prompt.Value);
            }
        }

        public static List<List<UserDefinedCellCells>> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var query = UserDefinedCellCells.lazy_query.Value;
            return query.GetFormulas(page, shapeids);
        }

        public static List<List<UserDefinedCellCells>> GetResults(IVisio.Page page, IList<int> shapeids)
        {
            var query = UserDefinedCellCells.lazy_query.Value;
            return query.GetResults(page, shapeids);
        }

        public static List<UserDefinedCellCells> GetFormulas(IVisio.Shape shape)
        {
            var query = UserDefinedCellCells.lazy_query.Value;
            return query.GetFormulas(shape);
        }

        public static List<UserDefinedCellCells> GetResults(IVisio.Shape shape)
        {
            var query = UserDefinedCellCells.lazy_query.Value;
            return query.GetResults(shape);
        }

        private static readonly System.Lazy<UserDefinedCellCellsReader> lazy_query = new System.Lazy<UserDefinedCellCellsReader>();
    }
}
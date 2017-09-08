using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;


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
            return query.GetValues(page, shapeids, CellValueType.Formula);
        }

        public static List<List<UserDefinedCellCells>> GetResults(IVisio.Page page, IList<int> shapeids)
        {
            var query = UserDefinedCellCells.lazy_query.Value;
            return query.GetValues(page, shapeids, CellValueType.Result);
        }

        public static List<UserDefinedCellCells> GetFormulas(IVisio.Shape shape)
        {
            var query = UserDefinedCellCells.lazy_query.Value;
            return query.GetValues(shape, CellValueType.Formula);
        }

        public static List<UserDefinedCellCells> GetResults(IVisio.Shape shape)
        {
            var query = UserDefinedCellCells.lazy_query.Value;
            return query.GetValues(shape, CellValueType.Result);
        }

        private static readonly System.Lazy<UserDefinedCellCellsReader> lazy_query = new System.Lazy<UserDefinedCellCellsReader>();


        class UserDefinedCellCellsReader : ReaderMultiRow<UserDefinedCellCells>
        {
            public SectionQueryColumn Value { get; set; }
            public SectionQueryColumn Prompt { get; set; }

            public UserDefinedCellCellsReader()
            {
                var sec = this.query.SectionQueries.Add(IVisio.VisSectionIndices.visSectionUser);
                this.Value = sec.Columns.Add(SrcConstants.UserDefCellValue, nameof(SrcConstants.UserDefCellValue));
                this.Prompt = sec.Columns.Add(SrcConstants.UserDefCellPrompt, nameof(SrcConstants.UserDefCellPrompt));
            }

            public override UserDefinedCellCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<string> row)
            {
                var cells = new UserDefinedCellCells();
                cells.Value = row[this.Value];
                cells.Prompt = row[this.Prompt];
                return cells;
            }
        }
    }
}
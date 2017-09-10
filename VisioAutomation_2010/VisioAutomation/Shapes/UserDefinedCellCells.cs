using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;


namespace VisioAutomation.Shapes
{
    public class UserDefinedCellCells : CellGroupMultiRow
    {
        public CellValueLiteral Value { get; set; }
        public CellValueLiteral Prompt { get; set; }

        public UserDefinedCellCells()
        {
        }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.UserDefCellValue, this.Value);
                yield return SrcValuePair.Create(SrcConstants.UserDefCellPrompt, this.Prompt);
            }
        }

        public static List<List<UserDefinedCellCells>> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType cvt)
        {
            var query = lazy_query.Value;
            return query.GetValues(page, shapeids, cvt);
        }

        public static List<UserDefinedCellCells> GetCells(IVisio.Shape shape, CellValueType cvt)
        {
            var query = lazy_query.Value;
            return query.GetValues(shape, cvt);
        }

        private static readonly System.Lazy<UserDefinedCellCellsReader> lazy_query = new System.Lazy<UserDefinedCellCellsReader>();


        class UserDefinedCellCellsReader : ReaderMultiRow<UserDefinedCellCells>
        {
            public SectionQueryColumn Value { get; set; }
            public SectionQueryColumn Prompt { get; set; }

            public UserDefinedCellCellsReader()
            {
                var sec = this.query.SectionQueries.Add(IVisio.VisSectionIndices.visSectionUser);
                this.Value = sec.Columns.Add(SrcConstants.UserDefCellValue, nameof(this.Value));
                this.Prompt = sec.Columns.Add(SrcConstants.UserDefCellPrompt, nameof(this.Prompt));
            }

            public override UserDefinedCellCells CellDataToCellGroup(Utilities.ArraySegment<string> row)
            {
                var cells = new UserDefinedCellCells();
                cells.Value = row[this.Value];
                cells.Prompt = row[this.Prompt];
                return cells;
            }
        }
    }
}
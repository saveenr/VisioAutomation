using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;


namespace VisioAutomation.Shapes
{
    public class UserDefinedCellCells : CellGroupBase
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

        public static List<List<UserDefinedCellCells>> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var query = lazy_query.Value;
            return query.GetCells(page, shapeids, type);
        }

        public static List<UserDefinedCellCells> GetCells(IVisio.Shape shape, CellValueType type)
        {
            var query = lazy_query.Value;
            return query.GetCells(shape, type);
        }

        private static readonly System.Lazy<UserDefinedCellCellsReader> lazy_query = new System.Lazy<UserDefinedCellCellsReader>();

        public void EncodeValues()
        {
            this.Value = CellValueLiteral.EncodeValue(this.Value.Value);
            this.Prompt = CellValueLiteral.EncodeValue(this.Prompt.Value);
        }


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

            public override UserDefinedCellCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new UserDefinedCellCells();
                cells.Value = row[this.Value];
                cells.Prompt = row[this.Prompt];
                return cells;
            }
        }
    }
}
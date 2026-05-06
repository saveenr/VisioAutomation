using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellRecords;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioAutomation.Shapes
{
    public class UserDefinedCellCells : CellRecord
    {
        public Core.CellValue Formula { get; set; }
        public Core.CellValue Prompt { get; set; }

        // Renamed Formula 2026-05-06 (issue #144). The cell stores a Visio formula,
        // not a literal value; the old name suggested otherwise. The Obsolete shim
        // preserves source compatibility through the deprecation window.
        [System.Obsolete("Renamed to Formula. The cell stores a Visio formula, not a literal value. Use SetString or SetFormula to set values without manual encoding.")]
        public Core.CellValue Value
        {
            get { return this.Formula; }
            set { this.Formula = value; }
        }

        public UserDefinedCellCells()
        {
        }

        public override IEnumerable<CellMetadata> GetCellMetadata()
        {
            yield return this._create(nameof(this.Formula), Core.SrcConstants.UserDefCellValue, this.Formula);
            yield return this._create(nameof(this.Prompt), Core.SrcConstants.UserDefCellPrompt, this.Prompt);
        }

        // === Typed setters (issue #144) ===
        // Mirror of CustomPropertyCells: emit correctly-encoded formulas without
        // requiring the caller to think about Visio's formula grammar.

        public void SetString(string value)
        {
            this.Formula = Core.CellValue.EncodeValue(value, true);
        }

        public void SetFormula(string formula)
        {
            // Raw escape hatch: assign the formula verbatim. For callers who have
            // constructed a valid Visio formula themselves.
            this.Formula = formula;
        }

        public void EncodeValues()
        {
            this.Formula = Core.CellValue.EncodeValue(this.Formula.Value);
            this.Prompt = Core.CellValue.EncodeValue(this.Prompt.Value);
        }

        public static CellRecordsGroup<UserDefinedCellCells> GetCells(IVisio.Page page, Core.ShapeIDPairs shapeidpairs,
            Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsMultipleShapesMultipleRows(page, shapeidpairs, type);
        }

        public static CellRecords<UserDefinedCellCells> GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsSingleShapeMultipleRows(shape, type);
        }

        private static readonly System.Lazy<Builder> builder = new System.Lazy<Builder>();

        public static UserDefinedCellCells RowToRecord(VASS.Data.DataRow<string> row, VASS.Data.DataColumns cols)
        {
            var record = new UserDefinedCellCells();
            var getcellvalue = getvalfromrowfunc(row, cols);

            record.Formula = getcellvalue(nameof(Formula));
            record.Prompt = getcellvalue(nameof(Prompt));


            return record;
        }


        class Builder : CellRecordBuilderSectionQuery<UserDefinedCellCells>
        {
            public Builder() : base(UserDefinedCellCells.RowToRecord)
            {
            }


        }
    }
}

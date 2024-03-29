using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellRecords;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioAutomation.Shapes
{
    public class UserDefinedCellCells : CellRecord
    {
        public Core.CellValue Value { get; set; }
        public Core.CellValue Prompt { get; set; }

        public UserDefinedCellCells()
        {
        }

        public override IEnumerable<CellMetadata> GetCellMetadata()
        {
            yield return this._create(nameof(this.Value), Core.SrcConstants.UserDefCellValue, this.Value);
            yield return this._create(nameof(this.Prompt), Core.SrcConstants.UserDefCellPrompt, this.Prompt);
        }

        public void EncodeValues()
        {
            this.Value = Core.CellValue.EncodeValue(this.Value.Value);
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

            record.Value = getcellvalue(nameof(Value));
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
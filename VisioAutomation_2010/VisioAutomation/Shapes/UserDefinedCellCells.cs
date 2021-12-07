using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellRecords;
using VACG = VisioAutomation.ShapeSheet.CellGroups;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioAutomation.Shapes
{
    public class UserDefinedCellCells : VACG.CellRecord
    {
        public Core.CellValue Value { get; set; }
        public Core.CellValue Prompt { get; set; }

        public UserDefinedCellCells()
        {
        }

        public override IEnumerable<VACG.CellMetadata> GetCellMetadata()
        {
            yield return this._create(nameof(this.Value), Core.SrcConstants.UserDefCellValue, this.Value);
            yield return this._create(nameof(this.Prompt), Core.SrcConstants.UserDefCellPrompt, this.Prompt);
        }

        public void EncodeValues()
        {
            this.Value = Core.CellValue.EncodeValue(this.Value.Value);
            this.Prompt = Core.CellValue.EncodeValue(this.Prompt.Value);
        }

        public static List<List<UserDefinedCellCells>> GetCells(IVisio.Page page, Core.ShapeIDPairs shapeidpairs,
            Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsMultipleShapesMultipleRows(page, shapeidpairs, type);
        }

        public static List<UserDefinedCellCells> GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsSingleShapeMultipleRows(shape, type);
        }

        private static readonly System.Lazy<Builder> builder = new System.Lazy<Builder>();


        class Builder : CellRecordBuilder<UserDefinedCellCells>
        {
            public Builder() : base(VACG.CellRecordBuilderType.MultiRow)
            {
            }


            public override UserDefinedCellCells ToCellGroup(VASS.Data.DataRow<string> row, VASS.Data.DataColumnCollection cols)
            {
                var cells = new UserDefinedCellCells();
                var getcellvalue = queryrow_to_cellgroup(row, cols);

                cells.Value = getcellvalue(nameof(Value));
                cells.Prompt = getcellvalue(nameof(Prompt));


                return cells;
            }
        }
    }
}
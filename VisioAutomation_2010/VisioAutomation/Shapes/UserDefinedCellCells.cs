using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioAutomation.Shapes
{
    public class UserDefinedCellCells : VASS.CellGroups.CellGroup
    {
        public VASS.CellValue Value { get; set; }
        public VASS.CellValue Prompt { get; set; }

        public UserDefinedCellCells()
        {
        }

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.Value), VASS.SrcConstants.UserDefCellValue, this.Value);
            yield return this.Create(nameof(this.Prompt), VASS.SrcConstants.UserDefCellPrompt, this.Prompt);
        }

        public void EncodeValues()
        {
            this.Value = VASS.CellValue.EncodeValue(this.Value.Value);
            this.Prompt = VASS.CellValue.EncodeValue(this.Prompt.Value);
        }

        public static List<List<UserDefinedCellCells>> GetCells(IVisio.Page page, Core.ShapeIDPairs shapeidpairs, VASS.CellValueType type)
        {
            var reader = UserDefinedCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(page, shapeidpairs, type);
        }

        public static List<UserDefinedCellCells> GetCells(IVisio.Shape shape, VASS.CellValueType type)
        {
            var reader = UserDefinedCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(shape, type);
        }

        private static readonly System.Lazy<UserDefinedCellCellsBuilder> UserDefinedCells_lazy_builder = new System.Lazy<UserDefinedCellCellsBuilder>();




        class UserDefinedCellCellsBuilder : VASS.CellGroups.CellGroupBuilder<UserDefinedCellCells>
        {

            public UserDefinedCellCellsBuilder() : base(VASS.CellGroups.CellGroupBuilderType.MultiRow)
            {
            }


            public override UserDefinedCellCells ToCellGroup(ShapeSheet.Query.Row<string> row, VisioAutomation.ShapeSheet.Query.Columns cols)
            {
                var cells = new UserDefinedCellCells();
                var getcellvalue = VisioAutomation.ShapeSheet.CellGroups.CellGroup.row_to_cellgroup(row, cols);

                cells.Value = getcellvalue(nameof(UserDefinedCellCells.Value));
                cells.Prompt = getcellvalue(nameof(UserDefinedCellCells.Prompt));



                return cells;
            }
        }

    }
}
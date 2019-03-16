using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioAutomation.Shapes
{
    public class UserDefinedCellCells : CellGroup
    {
        public VASS.CellValueLiteral Value { get; set; }
        public VASS.CellValueLiteral Prompt { get; set; }

        public UserDefinedCellCells()
        {
        }

        public override IEnumerable<CellMetadataItem> CellMetadata
        {
            get
            {


                yield return CellMetadataItem.Create(nameof(this.Value), VASS.SrcConstants.UserDefCellValue, this.Value);
                yield return CellMetadataItem.Create(nameof(this.Prompt), VASS.SrcConstants.UserDefCellPrompt, this.Prompt);
            }
        }

        public void EncodeValues()
        {
            this.Value = VASS.CellValueLiteral.EncodeValue(this.Value.Value);
            this.Prompt = VASS.CellValueLiteral.EncodeValue(this.Prompt.Value);
        }

        public static List<List<UserDefinedCellCells>> GetCells(IVisio.Page page, IList<int> shapeids, VASS.CellValueType type)
        {
            var reader = UserDefinedCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(page, shapeids, type);
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


            public override UserDefinedCellCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row, VisioAutomation.ShapeSheet.Query.ColumnList cols)
            {
                var cells = new UserDefinedCellCells();

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }


                cells.Value = getcellvalue(nameof(UserDefinedCellCells.Value));
                cells.Prompt = getcellvalue(nameof(UserDefinedCellCells.Prompt));



                return cells;
            }
        }

    }
}
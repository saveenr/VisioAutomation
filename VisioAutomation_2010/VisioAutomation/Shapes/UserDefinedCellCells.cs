using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Shapes
{
    public class UserDefinedCellCells : CellGroup
    {
        public CellValueLiteral Value { get; set; }
        public CellValueLiteral Prompt { get; set; }

        public UserDefinedCellCells()
        {
        }

        public override IEnumerable<CellMetadataItem> CellMetadata
        {
            get
            {


                yield return CellMetadataItem.Create(nameof(this.Value), SrcConstants.UserDefCellValue, this.Value);
                yield return CellMetadataItem.Create(nameof(this.Prompt), SrcConstants.UserDefCellPrompt, this.Prompt);
            }
        }

        public void EncodeValues()
        {
            this.Value = CellValueLiteral.EncodeValue(this.Value.Value);
            this.Prompt = CellValueLiteral.EncodeValue(this.Prompt.Value);
        }
    }
}
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

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.UserDefCellValue, this.Value);
                yield return SrcValuePair.Create(SrcConstants.UserDefCellPrompt, this.Prompt);
            }
        }

        public void EncodeValues()
        {
            this.Value = CellValueLiteral.EncodeValue(this.Value.Value);
            this.Prompt = CellValueLiteral.EncodeValue(this.Prompt.Value);
        }
    }
}
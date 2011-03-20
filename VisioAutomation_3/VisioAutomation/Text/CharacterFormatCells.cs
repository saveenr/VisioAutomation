using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Text
{
    public class CharacterFormatCells
    {
        public VA.ShapeSheet.CellData<int> Color { get; set; }
        public VA.ShapeSheet.CellData<int> Font { get; set; }
        public VA.ShapeSheet.CellData<double> Size { get; set; }
        public VA.ShapeSheet.CellData<VA.Text.CharStyle> Style { get; set; }
        public VA.ShapeSheet.CellData<double> Transparency { get; set; }

        public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, int row, short id)
        {
            short srow = (short) row;
            this._Apply((src, f) => update.SetFormulaIgnoreNull(id, src.ForRow(srow), f));
        }

        public void Apply(VA.ShapeSheet.Update.SRCUpdate update, int row)
        {
            short srow = (short) row;
            this._Apply((src, f) => update.SetFormulaIgnoreNull(src.ForRow(srow), f));
        }

        public void _Apply(System.Action<VA.ShapeSheet.SRC, VA.ShapeSheet.FormulaLiteral> func)
        {
            func(VA.ShapeSheet.SRCConstants.Char_Color, this.Color.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_Font, this.Font.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_Size, this.Size.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_Style, this.Style.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_ColorTrans, this.Transparency.Formula);
        }
    }
}
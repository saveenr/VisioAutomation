using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Text
{
    public class ParagraphFormatCells
    {
        ////public string BulletString;
        public VA.ShapeSheet.CellData<double> IndentFirst { get; set; }
        public VA.ShapeSheet.CellData<double> IndentRight { get; set; }
        public VA.ShapeSheet.CellData<double> IndentLeft { get; set; }
        public VA.ShapeSheet.CellData<double> SpacingBefore { get; set; }
        public VA.ShapeSheet.CellData<double> SpacingAfter { get; set; }
        public VA.ShapeSheet.CellData<double> SpacingLine { get; set; }
        public VA.ShapeSheet.CellData<int> HorizontalAlign { get; set; }
        public VA.ShapeSheet.CellData<int> BulletIndex { get; set; }
        public VA.ShapeSheet.CellData<int> BulletFont { get; set; }
        public VA.ShapeSheet.CellData<int> BulletSize { get; set; }

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

        private void _Apply(System.Action<VA.ShapeSheet.SRC, VA.ShapeSheet.FormulaLiteral> func)
        {
            func(VA.ShapeSheet.SRCConstants.Para_IndLeft, this.IndentLeft.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_IndFirst, this.IndentFirst.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_IndRight, this.IndentRight.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_SpAfter, this.SpacingAfter.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_SpBefore, this.SpacingBefore.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_SpLine, this.SpacingLine.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_HAlign, this.HorizontalAlign.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_BulletFont, this.BulletFont.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_BulletIndex, this.BulletIndex.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_BulletSize, this.BulletSize.Formula);
        }
    }
}
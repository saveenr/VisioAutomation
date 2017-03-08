using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Models.Text
{
    public class ParagraphFormatting 
    {
        public ShapeSheet.CellValueLiteral Bullet { get; set; }
        public ShapeSheet.CellValueLiteral BulletFont { get; set; }
        public ShapeSheet.CellValueLiteral BulletFontSize { get; set; }
        public ShapeSheet.CellValueLiteral BulletString { get; set; }
        public ShapeSheet.CellValueLiteral Flags { get; set; }
        public ShapeSheet.CellValueLiteral HorizontalAlign { get; set; }
        public ShapeSheet.CellValueLiteral IndentFirst { get; set; }
        public ShapeSheet.CellValueLiteral IndentLeft { get; set; }
        public ShapeSheet.CellValueLiteral IndentRight { get; set; }
        public ShapeSheet.CellValueLiteral LocBulletFont { get; set; }
        public ShapeSheet.CellValueLiteral SpacingAfter { get; set; }
        public ShapeSheet.CellValueLiteral SpacingBefore { get; set; }
        public ShapeSheet.CellValueLiteral SpacingLine { get; set; }
        public ShapeSheet.CellValueLiteral TextPosAfterBullet { get; set; }
        
        internal void ApplyFormulas(VisioAutomation.ShapeSheet.Writers.SrcWriter writer, short row)
        {
            writer.SetFormula(SrcConstants.ParaIndentLeft.CloneWithNewRow(row), this.IndentLeft);
            writer.SetFormula(SrcConstants.ParaIndentFirst.CloneWithNewRow(row), this.IndentFirst);
            writer.SetFormula(SrcConstants.ParaIndentRight.CloneWithNewRow(row), this.IndentRight);
            writer.SetFormula(SrcConstants.ParaSpacingAfter.CloneWithNewRow(row), this.SpacingAfter);
            writer.SetFormula(SrcConstants.ParaSpacingBefore.CloneWithNewRow(row), this.SpacingBefore);
            writer.SetFormula(SrcConstants.ParaSpacingLine.CloneWithNewRow(row), this.SpacingLine);
            writer.SetFormula(SrcConstants.ParaHorizontalAlign.CloneWithNewRow(row), this.HorizontalAlign);
            writer.SetFormula(SrcConstants.ParaBulletFont.CloneWithNewRow(row), this.BulletFont);
            writer.SetFormula(SrcConstants.ParaBullet.CloneWithNewRow(row), this.Bullet);
            writer.SetFormula(SrcConstants.ParaBulletFontSize.CloneWithNewRow(row), this.BulletFontSize);
            writer.SetFormula(SrcConstants.ParaLocalizeBulletFont.CloneWithNewRow(row), this.LocBulletFont);
            writer.SetFormula(SrcConstants.ParaTextPosAfterBullet.CloneWithNewRow(row), this.TextPosAfterBullet);
            writer.SetFormula(SrcConstants.ParaFlags.CloneWithNewRow(row), this.Flags);
            writer.SetFormula(SrcConstants.ParaBulletString.CloneWithNewRow(row), this.BulletString);
        }

        public void ApplyFormulasTo(ParagraphFormatting target)
        {
            if (this.IndentFirst.HasValue) { target.IndentFirst = this.IndentFirst; }
            if (this.IndentLeft.HasValue) { target.IndentLeft = this.IndentLeft; }
            if (this.IndentRight.HasValue) { target.IndentRight = this.IndentRight; }

            if (this.SpacingAfter.HasValue) { target.SpacingAfter = this.SpacingAfter; }
            if (this.SpacingBefore.HasValue) { target.SpacingBefore = this.SpacingBefore; }
            if (this.SpacingBefore.HasValue) { target.SpacingBefore = this.SpacingBefore; }

            if (this.Bullet.HasValue) { target.Bullet = this.Bullet; }
            if (this.BulletFont.HasValue) { target.BulletFont = this.BulletFont; }
            if (this.BulletFontSize.HasValue) { target.BulletFontSize = this.BulletFontSize; }

            if (this.BulletString.HasValue) { target.BulletString = this.BulletString; }
            if (this.Flags.HasValue) { target.Flags = this.Flags; }
            if (this.HorizontalAlign.HasValue) { target.HorizontalAlign = this.HorizontalAlign; }
            if (this.LocBulletFont.HasValue) { target.LocBulletFont = this.LocBulletFont; }
            if (this.TextPosAfterBullet.HasValue) { target.TextPosAfterBullet = this.TextPosAfterBullet; }
        }
    }
}
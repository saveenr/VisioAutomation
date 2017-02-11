using VisioAutomation.ShapeSheet.Writers;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.Models.Text
{
    public class ParagraphFormatting 
    {
        public ShapeSheet.FormulaLiteral Bullet { get; set; }
        public ShapeSheet.FormulaLiteral BulletFont { get; set; }
        public ShapeSheet.FormulaLiteral BulletFontSize { get; set; }
        public ShapeSheet.FormulaLiteral BulletString { get; set; }
        public ShapeSheet.FormulaLiteral Flags { get; set; }
        public ShapeSheet.FormulaLiteral HorizontalAlign { get; set; }
        public ShapeSheet.FormulaLiteral IndentFirst { get; set; }
        public ShapeSheet.FormulaLiteral IndentLeft { get; set; }
        public ShapeSheet.FormulaLiteral IndentRight { get; set; }
        public ShapeSheet.FormulaLiteral LocBulletFont { get; set; }
        public ShapeSheet.FormulaLiteral SpacingAfter { get; set; }
        public ShapeSheet.FormulaLiteral SpacingBefore { get; set; }
        public ShapeSheet.FormulaLiteral SpacingLine { get; set; }
        public ShapeSheet.FormulaLiteral TextPosAfterBullet { get; set; }
        
        internal void ApplyFormulas(ShapeSheetWriter writer, short row)
        {
            writer.SetFormula(SRCCON.Para_IndLeft.CloneWithNewRow(row), this.IndentLeft);
            writer.SetFormula(SRCCON.Para_IndFirst.CloneWithNewRow(row), this.IndentFirst);
            writer.SetFormula(SRCCON.Para_IndRight.CloneWithNewRow(row), this.IndentRight);
            writer.SetFormula(SRCCON.Para_SpAfter.CloneWithNewRow(row), this.SpacingAfter);
            writer.SetFormula(SRCCON.Para_SpBefore.CloneWithNewRow(row), this.SpacingBefore);
            writer.SetFormula(SRCCON.Para_SpLine.CloneWithNewRow(row), this.SpacingLine);
            writer.SetFormula(SRCCON.Para_HorzAlign.CloneWithNewRow(row), this.HorizontalAlign);
            writer.SetFormula(SRCCON.Para_BulletFont.CloneWithNewRow(row), this.BulletFont);
            writer.SetFormula(SRCCON.Para_Bullet.CloneWithNewRow(row), this.Bullet);
            writer.SetFormula(SRCCON.Para_BulletFontSize.CloneWithNewRow(row), this.BulletFontSize);
            writer.SetFormula(SRCCON.Para_LocalizeBulletFont.CloneWithNewRow(row), this.LocBulletFont);
            writer.SetFormula(SRCCON.Para_TextPosAfterBullet.CloneWithNewRow(row), this.TextPosAfterBullet);
            writer.SetFormula(SRCCON.Para_Flags.CloneWithNewRow(row), this.Flags);
            writer.SetFormula(SRCCON.Para_BulletStr.CloneWithNewRow(row), this.BulletString);
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
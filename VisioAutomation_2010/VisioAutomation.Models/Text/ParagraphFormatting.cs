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
        
        internal void ApplyFormulas(FormulaWriterSRC update, short row)
        {
            update.SetFormula(SRCCON.Para_IndLeft.WithRow(row), this.IndentLeft);
            update.SetFormula(SRCCON.Para_IndFirst.WithRow(row), this.IndentFirst);
            update.SetFormula(SRCCON.Para_IndRight.WithRow(row), this.IndentRight);
            update.SetFormula(SRCCON.Para_SpAfter.WithRow(row), this.SpacingAfter);
            update.SetFormula(SRCCON.Para_SpBefore.WithRow(row), this.SpacingBefore);
            update.SetFormula(SRCCON.Para_SpLine.WithRow(row), this.SpacingLine);
            update.SetFormula(SRCCON.Para_HorzAlign.WithRow(row), this.HorizontalAlign);
            update.SetFormula(SRCCON.Para_BulletFont.WithRow(row), this.BulletFont);
            update.SetFormula(SRCCON.Para_Bullet.WithRow(row), this.Bullet);
            update.SetFormula(SRCCON.Para_BulletFontSize.WithRow(row), this.BulletFontSize);
            update.SetFormula(SRCCON.Para_LocalizeBulletFont.WithRow(row), this.LocBulletFont);
            update.SetFormula(SRCCON.Para_TextPosAfterBullet.WithRow(row), this.TextPosAfterBullet);
            update.SetFormula(SRCCON.Para_Flags.WithRow(row), this.Flags);
            update.SetFormula(SRCCON.Para_BulletStr.WithRow(row), this.BulletString);
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
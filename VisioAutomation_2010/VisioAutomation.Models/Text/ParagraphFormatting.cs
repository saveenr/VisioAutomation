using VisioAutomation.ShapeSheet.Update;
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
        
        internal void ApplyFormulas(Update update, short row)
        {
            update.SetFormulaIgnoreNull(SRCCON.Para_IndLeft.ForRow(row), this.IndentLeft);
            update.SetFormulaIgnoreNull(SRCCON.Para_IndFirst.ForRow(row), this.IndentFirst);
            update.SetFormulaIgnoreNull(SRCCON.Para_IndRight.ForRow(row), this.IndentRight);
            update.SetFormulaIgnoreNull(SRCCON.Para_SpAfter.ForRow(row), this.SpacingAfter);
            update.SetFormulaIgnoreNull(SRCCON.Para_SpBefore.ForRow(row), this.SpacingBefore);
            update.SetFormulaIgnoreNull(SRCCON.Para_SpLine.ForRow(row), this.SpacingLine);
            update.SetFormulaIgnoreNull(SRCCON.Para_HorzAlign.ForRow(row), this.HorizontalAlign);
            update.SetFormulaIgnoreNull(SRCCON.Para_BulletFont.ForRow(row), this.BulletFont);
            update.SetFormulaIgnoreNull(SRCCON.Para_Bullet.ForRow(row), this.Bullet);
            update.SetFormulaIgnoreNull(SRCCON.Para_BulletFontSize.ForRow(row), this.BulletFontSize);
            update.SetFormulaIgnoreNull(SRCCON.Para_LocalizeBulletFont.ForRow(row), this.LocBulletFont);
            update.SetFormulaIgnoreNull(SRCCON.Para_TextPosAfterBullet.ForRow(row), this.TextPosAfterBullet);
            update.SetFormulaIgnoreNull(SRCCON.Para_Flags.ForRow(row), this.Flags);
            update.SetFormulaIgnoreNull(SRCCON.Para_BulletStr.ForRow(row), this.BulletString);
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
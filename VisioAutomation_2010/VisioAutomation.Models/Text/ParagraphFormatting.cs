namespace VisioAutomation.Models.Text
{
    public class ParagraphFormatting 
    {
        public Core.CellValue Bullet { get; set; }
        public Core.CellValue BulletFont { get; set; }
        public Core.CellValue BulletFontSize { get; set; }
        public Core.CellValue BulletString { get; set; }
        public Core.CellValue Flags { get; set; }
        public Core.CellValue HorizontalAlign { get; set; }
        public Core.CellValue IndentFirst { get; set; }
        public Core.CellValue IndentLeft { get; set; }
        public Core.CellValue IndentRight { get; set; }
        public Core.CellValue LocBulletFont { get; set; }
        public Core.CellValue SpacingAfter { get; set; }
        public Core.CellValue SpacingBefore { get; set; }
        public Core.CellValue SpacingLine { get; set; }
        public Core.CellValue TextPosAfterBullet { get; set; }
        
        internal void ApplyFormulas(VisioAutomation.ShapeSheet.Writers.SrcWriter writer, short row)
        {
            writer.SetValue(VisioAutomation.Core.SrcConstants.ParaIndentLeft.CloneWithNewRow(row), this.IndentLeft);
            writer.SetValue(VisioAutomation.Core.SrcConstants.ParaIndentFirst.CloneWithNewRow(row), this.IndentFirst);
            writer.SetValue(VisioAutomation.Core.SrcConstants.ParaIndentRight.CloneWithNewRow(row), this.IndentRight);
            writer.SetValue(VisioAutomation.Core.SrcConstants.ParaSpacingAfter.CloneWithNewRow(row), this.SpacingAfter);
            writer.SetValue(VisioAutomation.Core.SrcConstants.ParaSpacingBefore.CloneWithNewRow(row), this.SpacingBefore);
            writer.SetValue(VisioAutomation.Core.SrcConstants.ParaSpacingLine.CloneWithNewRow(row), this.SpacingLine);
            writer.SetValue(VisioAutomation.Core.SrcConstants.ParaHorizontalAlign.CloneWithNewRow(row), this.HorizontalAlign);
            writer.SetValue(VisioAutomation.Core.SrcConstants.ParaBulletFont.CloneWithNewRow(row), this.BulletFont);
            writer.SetValue(VisioAutomation.Core.SrcConstants.ParaBullet.CloneWithNewRow(row), this.Bullet);
            writer.SetValue(VisioAutomation.Core.SrcConstants.ParaBulletFontSize.CloneWithNewRow(row), this.BulletFontSize);
            writer.SetValue(VisioAutomation.Core.SrcConstants.ParaLocalizeBulletFont.CloneWithNewRow(row), this.LocBulletFont);
            writer.SetValue(VisioAutomation.Core.SrcConstants.ParaTextPosAfterBullet.CloneWithNewRow(row), this.TextPosAfterBullet);
            writer.SetValue(VisioAutomation.Core.SrcConstants.ParaFlags.CloneWithNewRow(row), this.Flags);
            writer.SetValue(VisioAutomation.Core.SrcConstants.ParaBulletString.CloneWithNewRow(row), this.BulletString);
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
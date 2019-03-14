using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Text
{
    public class ParagraphFormatCells : CellGroup
    {
        public CellValueLiteral IndentFirst { get; set; }
        public CellValueLiteral IndentRight { get; set; }
        public CellValueLiteral IndentLeft { get; set; }
        public CellValueLiteral SpacingBefore { get; set; }
        public CellValueLiteral SpacingAfter { get; set; }
        public CellValueLiteral SpacingLine { get; set; }
        public CellValueLiteral HorizontalAlign { get; set; }
        public CellValueLiteral Bullet { get; set; }
        public CellValueLiteral BulletFont { get; set; }
        public CellValueLiteral BulletFontSize { get; set; }
        public CellValueLiteral LocalizeBulletFont { get; set; }
        public CellValueLiteral TextPosAfterBullet { get; set; }
        public CellValueLiteral Flags { get; set; }
        public CellValueLiteral BulletString { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.ParaIndentLeft, this.IndentLeft);
                yield return SrcValuePair.Create(SrcConstants.ParaIndentFirst, this.IndentFirst);
                yield return SrcValuePair.Create(SrcConstants.ParaIndentRight, this.IndentRight);
                yield return SrcValuePair.Create(SrcConstants.ParaSpacingAfter, this.SpacingAfter);
                yield return SrcValuePair.Create(SrcConstants.ParaSpacingBefore, this.SpacingBefore);
                yield return SrcValuePair.Create(SrcConstants.ParaSpacingLine, this.SpacingLine);
                yield return SrcValuePair.Create(SrcConstants.ParaHorizontalAlign, this.HorizontalAlign);
                yield return SrcValuePair.Create(SrcConstants.ParaBulletFont, this.BulletFont);
                yield return SrcValuePair.Create(SrcConstants.ParaBullet, this.Bullet);
                yield return SrcValuePair.Create(SrcConstants.ParaBulletFontSize, this.BulletFontSize);
                yield return SrcValuePair.Create(SrcConstants.ParaLocalizeBulletFont, this.LocalizeBulletFont);
                yield return SrcValuePair.Create(SrcConstants.ParaTextPosAfterBullet, this.TextPosAfterBullet);
                yield return SrcValuePair.Create(SrcConstants.ParaFlags, this.Flags);
                yield return SrcValuePair.Create(SrcConstants.ParaBulletString, this.BulletString);
            }
        }

    }
} 
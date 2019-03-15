using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;

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

        public override IEnumerable<CellMetadataItem> CellMetadata
        {
            get
            {


                yield return CellMetadataItem.Create(nameof(this.IndentLeft), SrcConstants.ParaIndentLeft, this.IndentLeft);
                yield return CellMetadataItem.Create(nameof(this.IndentFirst), SrcConstants.ParaIndentFirst, this.IndentFirst);
                yield return CellMetadataItem.Create(nameof(this.IndentRight), SrcConstants.ParaIndentRight, this.IndentRight);
                yield return CellMetadataItem.Create(nameof(this.SpacingAfter), SrcConstants.ParaSpacingAfter, this.SpacingAfter);
                yield return CellMetadataItem.Create(nameof(this.SpacingBefore), SrcConstants.ParaSpacingBefore, this.SpacingBefore);
                yield return CellMetadataItem.Create(nameof(this.SpacingLine), SrcConstants.ParaSpacingLine, this.SpacingLine);
                yield return CellMetadataItem.Create(nameof(this.HorizontalAlign), SrcConstants.ParaHorizontalAlign, this.HorizontalAlign);
                yield return CellMetadataItem.Create(nameof(this.BulletFont), SrcConstants.ParaBulletFont, this.BulletFont);
                yield return CellMetadataItem.Create(nameof(this.Bullet), SrcConstants.ParaBullet, this.Bullet);
                yield return CellMetadataItem.Create(nameof(this.BulletFontSize), SrcConstants.ParaBulletFontSize, this.BulletFontSize);
                yield return CellMetadataItem.Create(nameof(this.LocalizeBulletFont), SrcConstants.ParaLocalizeBulletFont, this.LocalizeBulletFont);
                yield return CellMetadataItem.Create(nameof(this.TextPosAfterBullet), SrcConstants.ParaTextPosAfterBullet, this.TextPosAfterBullet);
                yield return CellMetadataItem.Create(nameof(this.Flags), SrcConstants.ParaFlags, this.Flags);
                yield return CellMetadataItem.Create(nameof(this.BulletString), SrcConstants.ParaBulletString, this.BulletString);
            }
        }

    }
} 
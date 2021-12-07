using System.Collections.Generic;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;
using VACG=VisioAutomation.ShapeSheet.CellGroups;

namespace VisioAutomation.Text
{
    public class ParagraphCells : VACG.CellRecord
    {
        public Core.CellValue IndentFirst { get; set; }
        public Core.CellValue IndentRight { get; set; }
        public Core.CellValue IndentLeft { get; set; }
        public Core.CellValue SpacingBefore { get; set; }
        public Core.CellValue SpacingAfter { get; set; }
        public Core.CellValue SpacingLine { get; set; }
        public Core.CellValue HorizontalAlign { get; set; }
        public Core.CellValue Bullet { get; set; }
        public Core.CellValue BulletFont { get; set; }
        public Core.CellValue BulletFontSize { get; set; }
        public Core.CellValue LocalizeBulletFont { get; set; }
        public Core.CellValue TextPosAfterBullet { get; set; }
        public Core.CellValue Flags { get; set; }
        public Core.CellValue BulletString { get; set; }

        public override IEnumerable<VACG.CellMetadata> GetCellMetadata()
        {
            yield return this._create(nameof(this.IndentLeft), Core.SrcConstants.ParaIndentLeft, this.IndentLeft);
            yield return this._create(nameof(this.IndentFirst), Core.SrcConstants.ParaIndentFirst, this.IndentFirst);
            yield return this._create(nameof(this.IndentRight), Core.SrcConstants.ParaIndentRight, this.IndentRight);
            yield return this._create(nameof(this.SpacingAfter), Core.SrcConstants.ParaSpacingAfter, this.SpacingAfter);
            yield return this._create(nameof(this.SpacingBefore), Core.SrcConstants.ParaSpacingBefore, this.SpacingBefore);
            yield return this._create(nameof(this.SpacingLine), Core.SrcConstants.ParaSpacingLine, this.SpacingLine);
            yield return this._create(nameof(this.HorizontalAlign), Core.SrcConstants.ParaHorizontalAlign, this.HorizontalAlign);
            yield return this._create(nameof(this.BulletFont), Core.SrcConstants.ParaBulletFont, this.BulletFont);
            yield return this._create(nameof(this.Bullet), Core.SrcConstants.ParaBullet, this.Bullet);
            yield return this._create(nameof(this.BulletFontSize), Core.SrcConstants.ParaBulletFontSize, this.BulletFontSize);
            yield return this._create(nameof(this.LocalizeBulletFont), Core.SrcConstants.ParaLocalizeBulletFont,
                this.LocalizeBulletFont);
            yield return this._create(nameof(this.TextPosAfterBullet), Core.SrcConstants.ParaTextPosAfterBullet,
                this.TextPosAfterBullet);
            yield return this._create(nameof(this.Flags), Core.SrcConstants.ParaFlags, this.Flags);
            yield return this._create(nameof(this.BulletString), Core.SrcConstants.ParaBulletString, this.BulletString);
        }

        public static List<List<ParagraphCells>> GetCells(IVisio.Page page, Core.ShapeIDPairs shapeidpairs, Core.CellValueType type)

        {
            var reader = builder.Value;
            return reader.GetCellsMultipleShapesMultipleRows(page, shapeidpairs, type);
        }

        public static List<ParagraphCells> GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsSingleShapeMultipleRows(shape, type);
        }


        private static readonly System.Lazy<Builder> builder = new System.Lazy<Builder>();


        class Builder : VACG.CellGroupBuilder<ParagraphCells>
        {
            public Builder() : base(VACG.CellGroupBuilderType.MultiRow)
            {
            }

            public override ParagraphCells ToCellGroup(VASS.Data.DataRow<string> row, VASS.Data.DataColumnCollection cols)
            {
                var getcellvalue = queryrow_to_cellgroup(row, cols);
                var cells = new ParagraphCells();

                cells.IndentFirst = getcellvalue(nameof(IndentFirst));
                cells.IndentLeft = getcellvalue(nameof(IndentLeft));
                cells.IndentRight = getcellvalue(nameof(IndentRight));
                cells.SpacingAfter = getcellvalue(nameof(SpacingAfter));
                cells.SpacingBefore = getcellvalue(nameof(SpacingBefore));
                cells.SpacingLine = getcellvalue(nameof(SpacingLine));
                cells.HorizontalAlign = getcellvalue(nameof(HorizontalAlign));
                cells.Bullet = getcellvalue(nameof(Bullet));
                cells.BulletFont = getcellvalue(nameof(BulletFont));
                cells.BulletFontSize = getcellvalue(nameof(BulletFontSize));
                cells.LocalizeBulletFont = getcellvalue(nameof(LocalizeBulletFont));
                cells.TextPosAfterBullet = getcellvalue(nameof(TextPosAfterBullet));
                cells.Flags = getcellvalue(nameof(Flags));
                cells.BulletString = getcellvalue(nameof(BulletString));

                return cells;
            }
        }

    }
} 
using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

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


                yield return this.Create(nameof(this.IndentLeft), SrcConstants.ParaIndentLeft, this.IndentLeft);
                yield return this.Create(nameof(this.IndentFirst), SrcConstants.ParaIndentFirst, this.IndentFirst);
                yield return this.Create(nameof(this.IndentRight), SrcConstants.ParaIndentRight, this.IndentRight);
                yield return this.Create(nameof(this.SpacingAfter), SrcConstants.ParaSpacingAfter, this.SpacingAfter);
                yield return this.Create(nameof(this.SpacingBefore), SrcConstants.ParaSpacingBefore, this.SpacingBefore);
                yield return this.Create(nameof(this.SpacingLine), SrcConstants.ParaSpacingLine, this.SpacingLine);
                yield return this.Create(nameof(this.HorizontalAlign), SrcConstants.ParaHorizontalAlign, this.HorizontalAlign);
                yield return this.Create(nameof(this.BulletFont), SrcConstants.ParaBulletFont, this.BulletFont);
                yield return this.Create(nameof(this.Bullet), SrcConstants.ParaBullet, this.Bullet);
                yield return this.Create(nameof(this.BulletFontSize), SrcConstants.ParaBulletFontSize, this.BulletFontSize);
                yield return this.Create(nameof(this.LocalizeBulletFont), SrcConstants.ParaLocalizeBulletFont, this.LocalizeBulletFont);
                yield return this.Create(nameof(this.TextPosAfterBullet), SrcConstants.ParaTextPosAfterBullet, this.TextPosAfterBullet);
                yield return this.Create(nameof(this.Flags), SrcConstants.ParaFlags, this.Flags);
                yield return this.Create(nameof(this.BulletString), SrcConstants.ParaBulletString, this.BulletString);
            }
        }

        public static List<List<ParagraphFormatCells>> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = ParagraphFormatCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(page, shapeids, type);
        }

        public static List<ParagraphFormatCells> GetCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = ParagraphFormatCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(shape, type);
        }


        private static readonly System.Lazy<ParagraphFormatCellsBuilder> ParagraphFormatCells_lazy_builder = new System.Lazy<ParagraphFormatCellsBuilder>();


        class ParagraphFormatCellsBuilder : CellGroupBuilder<Text.ParagraphFormatCells>
        {
            public ParagraphFormatCellsBuilder() : base(CellGroupBuilderType.MultiRow)
            {
            }

            public override Text.ParagraphFormatCells ToCellGroup(ShapeSheet.Query.Row<string> row, VisioAutomation.ShapeSheet.Query.Columns cols)
            {
                var getcellvalue = VisioAutomation.ShapeSheet.CellGroups.CellGroup.gcf(row, cols);

                var cells = new Text.ParagraphFormatCells();

                var names = cells.CellMetadata.Select(i => i.Name).ToList();



                cells.IndentFirst = getcellvalue(nameof(ParagraphFormatCells.IndentFirst));
                cells.IndentLeft = getcellvalue(nameof(ParagraphFormatCells.IndentLeft));
                cells.IndentRight = getcellvalue(nameof(ParagraphFormatCells.IndentRight));
                cells.SpacingAfter = getcellvalue(nameof(ParagraphFormatCells.SpacingAfter));
                cells.SpacingBefore = getcellvalue(nameof(ParagraphFormatCells.SpacingBefore));
                cells.SpacingLine = getcellvalue(nameof(ParagraphFormatCells.SpacingLine));
                cells.HorizontalAlign = getcellvalue(nameof(ParagraphFormatCells.HorizontalAlign));
                cells.Bullet = getcellvalue(nameof(ParagraphFormatCells.Bullet));
                cells.BulletFont = getcellvalue(nameof(ParagraphFormatCells.BulletFont));
                cells.BulletFontSize = getcellvalue(nameof(ParagraphFormatCells.BulletFontSize));
                cells.LocalizeBulletFont = getcellvalue(nameof(ParagraphFormatCells.LocalizeBulletFont));
                cells.TextPosAfterBullet = getcellvalue(nameof(ParagraphFormatCells.TextPosAfterBullet));
                cells.Flags = getcellvalue(nameof(ParagraphFormatCells.Flags));
                cells.BulletString = getcellvalue(nameof(ParagraphFormatCells.BulletString));

                return cells;
            }
        }

    }
} 
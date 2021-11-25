using System.Collections.Generic;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;
using VisioAutomation.ShapeSheet.CellGroups;

namespace VisioAutomation.Text
{
    public class ParagraphFormatCells : VASS.CellGroups.CellGroup
    {
        public VASS.CellValue IndentFirst { get; set; }
        public VASS.CellValue IndentRight { get; set; }
        public VASS.CellValue IndentLeft { get; set; }
        public VASS.CellValue SpacingBefore { get; set; }
        public VASS.CellValue SpacingAfter { get; set; }
        public VASS.CellValue SpacingLine { get; set; }
        public VASS.CellValue HorizontalAlign { get; set; }
        public VASS.CellValue Bullet { get; set; }
        public VASS.CellValue BulletFont { get; set; }
        public VASS.CellValue BulletFontSize { get; set; }
        public VASS.CellValue LocalizeBulletFont { get; set; }
        public VASS.CellValue TextPosAfterBullet { get; set; }
        public VASS.CellValue Flags { get; set; }
        public VASS.CellValue BulletString { get; set; }

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.IndentLeft), VASS.SrcConstants.ParaIndentLeft, this.IndentLeft);
            yield return this.Create(nameof(this.IndentFirst), VASS.SrcConstants.ParaIndentFirst, this.IndentFirst);
            yield return this.Create(nameof(this.IndentRight), VASS.SrcConstants.ParaIndentRight, this.IndentRight);
            yield return this.Create(nameof(this.SpacingAfter), VASS.SrcConstants.ParaSpacingAfter, this.SpacingAfter);
            yield return this.Create(nameof(this.SpacingBefore), VASS.SrcConstants.ParaSpacingBefore, this.SpacingBefore);
            yield return this.Create(nameof(this.SpacingLine), VASS.SrcConstants.ParaSpacingLine, this.SpacingLine);
            yield return this.Create(nameof(this.HorizontalAlign), VASS.SrcConstants.ParaHorizontalAlign, this.HorizontalAlign);
            yield return this.Create(nameof(this.BulletFont), VASS.SrcConstants.ParaBulletFont, this.BulletFont);
            yield return this.Create(nameof(this.Bullet), VASS.SrcConstants.ParaBullet, this.Bullet);
            yield return this.Create(nameof(this.BulletFontSize), VASS.SrcConstants.ParaBulletFontSize, this.BulletFontSize);
            yield return this.Create(nameof(this.LocalizeBulletFont), VASS.SrcConstants.ParaLocalizeBulletFont,
                this.LocalizeBulletFont);
            yield return this.Create(nameof(this.TextPosAfterBullet), VASS.SrcConstants.ParaTextPosAfterBullet,
                this.TextPosAfterBullet);
            yield return this.Create(nameof(this.Flags), VASS.SrcConstants.ParaFlags, this.Flags);
            yield return this.Create(nameof(this.BulletString), VASS.SrcConstants.ParaBulletString, this.BulletString);
        }

        public static List<List<ParagraphFormatCells>> GetCells(IVisio.Page page, ShapeIDPairs shapeidpairs, VASS.CellValueType type)

        {
            var reader = ParagraphFormatCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(page, shapeidpairs, type);
        }

        public static List<ParagraphFormatCells> GetCells(IVisio.Shape shape, VASS.CellValueType type)
        {
            var reader = ParagraphFormatCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(shape, type);
        }


        private static readonly System.Lazy<ParagraphFormatCellsBuilder> ParagraphFormatCells_lazy_builder = new System.Lazy<ParagraphFormatCellsBuilder>();


        class ParagraphFormatCellsBuilder : VASS.CellGroups.CellGroupBuilder<Text.ParagraphFormatCells>
        {
            public ParagraphFormatCellsBuilder() : base(VASS.CellGroups.CellGroupBuilderType.MultiRow)
            {
            }

            public override Text.ParagraphFormatCells ToCellGroup(ShapeSheet.Query.Row<string> row, VisioAutomation.ShapeSheet.Query.Columns cols)
            {
                var getcellvalue = VisioAutomation.ShapeSheet.CellGroups.CellGroup.row_to_cellgroup(row, cols);

                var cells = new Text.ParagraphFormatCells();

                var names = cells.GetCellMetadata().Select(i => i.Name).ToList();



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
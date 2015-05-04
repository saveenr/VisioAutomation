using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Text
{
    public class ParagraphCells : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        ////public string BulletString;
        public ShapeSheet.CellData<double> IndentFirst { get; set; }
        public ShapeSheet.CellData<double> IndentRight { get; set; }
        public ShapeSheet.CellData<double> IndentLeft { get; set; }
        public ShapeSheet.CellData<double> SpacingBefore { get; set; }
        public ShapeSheet.CellData<double> SpacingAfter { get; set; }
        public ShapeSheet.CellData<double> SpacingLine { get; set; }
        public ShapeSheet.CellData<int> HorizontalAlign { get; set; }
        public ShapeSheet.CellData<int> Bullet { get; set; }
        public ShapeSheet.CellData<int> BulletFont { get; set; }
        public ShapeSheet.CellData<int> BulletFontSize { get; set; }
        public ShapeSheet.CellData<int> LocBulletFont { get; set; }
        public ShapeSheet.CellData<double> TextPosAfterBullet { get; set; }
        public ShapeSheet.CellData<int> Flags { get; set; }
        public ShapeSheet.CellData<string> BulletString { get; set; }

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SRCConstants.Para_IndLeft, this.IndentLeft.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_IndFirst, this.IndentFirst.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_IndRight, this.IndentRight.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_SpAfter, this.SpacingAfter.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_SpBefore, this.SpacingBefore.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_SpLine, this.SpacingLine.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_HorzAlign, this.HorizontalAlign.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_BulletFont, this.BulletFont.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_Bullet, this.Bullet.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_BulletFontSize, this.BulletFontSize.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_LocalizeBulletFont, this.LocBulletFont.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_TextPosAfterBullet, this.TextPosAfterBullet.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_Flags, this.Flags.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Para_BulletStr, this.BulletString.Formula);
            }
        }

        public static IList<List<ParagraphCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = ParagraphCells.get_query();
            return CellGroupMultiRow._GetCells<ParagraphCells,double>(page, shapeids, query, query.GetCells);
        }

        public static IList<ParagraphCells> GetCells(IVisio.Shape shape)
        {
            var query = ParagraphCells.get_query();
            return CellGroupMultiRow._GetCells<ParagraphCells,double>(shape, query, query.GetCells);
        }

        private static ParagraphFormatCellQuery _mCellQuery;
        private static ParagraphFormatCellQuery get_query()
        {
            ParagraphCells._mCellQuery = ParagraphCells._mCellQuery ?? new ParagraphFormatCellQuery();
            return ParagraphCells._mCellQuery;
        }

        class ParagraphFormatCellQuery : CellQuery
        {
            public CellColumn Bullet { get; set; }
            public CellColumn BulletFont { get; set; }
            public CellColumn BulletFontSize { get; set; }
            public CellColumn BulletString { get; set; } // NOTE: This is never used
            public CellColumn Flags { get; set; }
            public CellColumn HorzAlign { get; set; }
            public CellColumn IndentFirst { get; set; }
            public CellColumn IndentLeft { get; set; }
            public CellColumn IndentRight { get; set; }
            public CellColumn LocalizeBulletFont { get; set; }
            public CellColumn SpaceAfter { get; set; }
            public CellColumn SpaceBefore { get; set; }
            public CellColumn SpaceLine { get; set; }
            public CellColumn TextPosAfterBullet { get; set; }

            public ParagraphFormatCellQuery() 
            {
                var sec = this.AddSection(IVisio.VisSectionIndices.visSectionParagraph);
                this.Bullet = sec.AddCell(ShapeSheet.SRCConstants.Para_Bullet, "Para_Bullet");
                this.BulletFont = sec.AddCell(ShapeSheet.SRCConstants.Para_BulletFont, "Para_BulletFont");
                this.BulletFontSize = sec.AddCell(ShapeSheet.SRCConstants.Para_BulletFontSize, "Para_BulletFontSize");
                this.BulletString = sec.AddCell(ShapeSheet.SRCConstants.Para_BulletStr, "Para_BulletStr");
                this.Flags = sec.AddCell(ShapeSheet.SRCConstants.Para_Flags, "Para_Flags");
                this.HorzAlign = sec.AddCell(ShapeSheet.SRCConstants.Para_HorzAlign, "Para_HorzAlign");
                this.IndentFirst = sec.AddCell(ShapeSheet.SRCConstants.Para_IndFirst, "Para_IndFirst");
                this.IndentLeft = sec.AddCell(ShapeSheet.SRCConstants.Para_IndLeft, "Para_IndLeft");
                this.IndentRight = sec.AddCell(ShapeSheet.SRCConstants.Para_IndRight, "Para_IndRight");
                this.LocalizeBulletFont = sec.AddCell(ShapeSheet.SRCConstants.Para_LocalizeBulletFont, "Para_LocalizeBulletFont");
                this.SpaceAfter = sec.AddCell(ShapeSheet.SRCConstants.Para_SpAfter, "Para_SpAfter");
                this.SpaceBefore = sec.AddCell(ShapeSheet.SRCConstants.Para_SpBefore, "Para_SpBefore");
                this.SpaceLine = sec.AddCell(ShapeSheet.SRCConstants.Para_SpLine, "Para_SpLine");
                this.TextPosAfterBullet = sec.AddCell(ShapeSheet.SRCConstants.Para_TextPosAfterBullet, "Para_TextPosAfterBullet");
            }

            public ParagraphCells GetCells(IList<ShapeSheet.CellData<double>> row)
            {
                var cells = new ParagraphCells();
                cells.IndentFirst = row[this.IndentFirst];
                cells.IndentLeft = row[this.IndentLeft];
                cells.IndentRight = row[this.IndentRight];
                cells.SpacingAfter = row[this.SpaceAfter];
                cells.SpacingBefore = row[this.SpaceBefore];
                cells.SpacingLine = row[this.SpaceLine];
                cells.HorizontalAlign = row[this.HorzAlign].ToInt();
                cells.Bullet = row[this.Bullet].ToInt();
                cells.BulletFont = row[this.BulletFont].ToInt();
                cells.BulletFontSize = row[this.BulletFontSize].ToInt();
                cells.LocBulletFont = row[this.LocalizeBulletFont].ToInt();
                cells.TextPosAfterBullet = row[this.TextPosAfterBullet];
                cells.Flags = row[this.Flags].ToInt();
                cells.BulletString = ""; // TODO: Figure out some way of getting this

                return cells;
            }
        }
    }
} 
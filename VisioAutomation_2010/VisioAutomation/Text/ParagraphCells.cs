using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Text
{
    public class ParagraphCells : VA.ShapeSheet.CellGroups.CellGroupMultiRow
    {
        ////public string BulletString;
        public VA.ShapeSheet.CellData<double> IndentFirst { get; set; }
        public VA.ShapeSheet.CellData<double> IndentRight { get; set; }
        public VA.ShapeSheet.CellData<double> IndentLeft { get; set; }
        public VA.ShapeSheet.CellData<double> SpacingBefore { get; set; }
        public VA.ShapeSheet.CellData<double> SpacingAfter { get; set; }
        public VA.ShapeSheet.CellData<double> SpacingLine { get; set; }
        public VA.ShapeSheet.CellData<int> HorizontalAlign { get; set; }
        public VA.ShapeSheet.CellData<int> Bullet { get; set; }
        public VA.ShapeSheet.CellData<int> BulletFont { get; set; }
        public VA.ShapeSheet.CellData<int> BulletFontSize { get; set; }
        public VA.ShapeSheet.CellData<int> LocBulletFont { get; set; }
        public VA.ShapeSheet.CellData<double> TextPosAfterBullet { get; set; }
        public VA.ShapeSheet.CellData<int> Flags { get; set; }
        public VA.ShapeSheet.CellData<string> BulletString { get; set; }

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return newpair(VA.ShapeSheet.SRCConstants.Para_IndLeft, this.IndentLeft.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Para_IndFirst, this.IndentFirst.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Para_IndRight, this.IndentRight.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Para_SpAfter, this.SpacingAfter.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Para_SpBefore, this.SpacingBefore.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Para_SpLine, this.SpacingLine.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Para_HorzAlign, this.HorizontalAlign.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Para_BulletFont, this.BulletFont.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Para_Bullet, this.Bullet.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Para_BulletFontSize, this.BulletFontSize.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Para_LocalizeBulletFont, this.LocBulletFont.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Para_TextPosAfterBullet, this.TextPosAfterBullet.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Para_Flags, this.Flags.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.Para_BulletStr, this.BulletString.Formula);
            }
        }

        public static IList<List<ParagraphCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return _GetCells<ParagraphCells,double>(page, shapeids, query, query.GetCells);
        }

        public static IList<ParagraphCells> GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return _GetCells<ParagraphCells,double>(shape, query, query.GetCells);
        }

        private static ParagraphFormatCellQuery _mCellQuery;
        private static ParagraphFormatCellQuery get_query()
        {
            _mCellQuery = _mCellQuery ?? new ParagraphFormatCellQuery();
            return _mCellQuery;
        }

        class ParagraphFormatCellQuery : VA.ShapeSheet.Query.CellQuery
        {
            public CellColumn Bullet { get; set; }
            public CellColumn BulletFont { get; set; }
            public CellColumn BulletFontSize { get; set; }
            public CellColumn BulletString { get; set; }
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
                Bullet = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_Bullet, "BulletIndex");
                BulletFont = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_BulletFont, "BulletFont");
                BulletFontSize = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_BulletFontSize, "BulletFontSize");
                BulletString = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_BulletStr, "BulletString");
                Flags = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_Flags, "Flags");
                HorzAlign = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_HorzAlign, "HorzAlign");
                IndentFirst = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_IndFirst, "IndentFirst");
                IndentLeft = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_IndLeft, "IndentLeft");
                IndentRight = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_IndRight, "IndentRight");
                LocalizeBulletFont = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_LocalizeBulletFont, "LocalizeBulletFont");
                SpaceAfter = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_SpAfter, "SpaceAfter");
                SpaceBefore = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_SpBefore, "SpaceBefore");
                SpaceLine = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_SpLine, "SpaceLine");
                TextPosAfterBullet = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_TextPosAfterBullet, "TextPosAfterBullet");
            }

            public ParagraphCells GetCells(IList<VA.ShapeSheet.CellData<double>> row)
            {
                var cells = new ParagraphCells();
                cells.IndentFirst = row[IndentFirst.Ordinal];
                cells.IndentLeft = row[IndentLeft.Ordinal];
                cells.IndentRight = row[IndentRight.Ordinal];
                cells.SpacingAfter = row[SpaceAfter.Ordinal];
                cells.SpacingBefore = row[SpaceBefore.Ordinal];
                cells.SpacingLine = row[SpaceLine.Ordinal];
                cells.HorizontalAlign = row[HorzAlign.Ordinal].ToInt();
                cells.Bullet = row[Bullet.Ordinal].ToInt();
                cells.BulletFont = row[BulletFont.Ordinal].ToInt();
                cells.BulletFontSize = row[BulletFontSize.Ordinal].ToInt();
                cells.LocBulletFont = row[LocalizeBulletFont.Ordinal].ToInt();
                cells.TextPosAfterBullet = row[TextPosAfterBullet.Ordinal];
                cells.Flags = row[Flags.Ordinal].ToInt();
                cells.BulletString = ""; // TODO: Figure out some way of getting this

                return cells;
            }
        }
    }
} 
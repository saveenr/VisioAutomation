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
                Bullet = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_Bullet, "Para_Bullet");
                BulletFont = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_BulletFont, "Para_BulletFont");
                BulletFontSize = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_BulletFontSize, "Para_BulletFontSize");
                BulletString = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_BulletStr, "Para_BulletStr");
                Flags = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_Flags, "Para_Flags");
                HorzAlign = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_HorzAlign, "Para_HorzAlign");
                IndentFirst = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_IndFirst, "Para_IndFirst");
                IndentLeft = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_IndLeft, "Para_IndLeft");
                IndentRight = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_IndRight, "Para_IndRight");
                LocalizeBulletFont = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_LocalizeBulletFont, "Para_LocalizeBulletFont");
                SpaceAfter = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_SpAfter, "Para_SpAfter");
                SpaceBefore = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_SpBefore, "Para_SpBefore");
                SpaceLine = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_SpLine, "Para_SpLine");
                TextPosAfterBullet = sec.AddCell(VA.ShapeSheet.SRCConstants.Para_TextPosAfterBullet, "Para_TextPosAfterBullet");
            }

            public ParagraphCells GetCells(IList<VA.ShapeSheet.CellData<double>> row)
            {
                var cells = new ParagraphCells();
                cells.IndentFirst = row[IndentFirst];
                cells.IndentLeft = row[IndentLeft];
                cells.IndentRight = row[IndentRight];
                cells.SpacingAfter = row[SpaceAfter];
                cells.SpacingBefore = row[SpaceBefore];
                cells.SpacingLine = row[SpaceLine];
                cells.HorizontalAlign = row[HorzAlign].ToInt();
                cells.Bullet = row[Bullet].ToInt();
                cells.BulletFont = row[BulletFont].ToInt();
                cells.BulletFontSize = row[BulletFontSize].ToInt();
                cells.LocBulletFont = row[LocalizeBulletFont].ToInt();
                cells.TextPosAfterBullet = row[TextPosAfterBullet];
                cells.Flags = row[Flags].ToInt();
                cells.BulletString = ""; // TODO: Figure out some way of getting this

                return cells;
            }
        }
    }
} 
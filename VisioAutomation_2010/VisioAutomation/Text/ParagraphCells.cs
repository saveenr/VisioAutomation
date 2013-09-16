using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

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

        public override void ApplyFormulasForRow(ApplyFormula func, short row)
        {
            func(VA.ShapeSheet.SRCConstants.Para_IndLeft.ForRow(row), this.IndentLeft.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_IndFirst.ForRow(row), this.IndentFirst.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_IndRight.ForRow(row), this.IndentRight.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_SpAfter.ForRow(row), this.SpacingAfter.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_SpBefore.ForRow(row), this.SpacingBefore.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_SpLine.ForRow(row), this.SpacingLine.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_HorzAlign.ForRow(row), this.HorizontalAlign.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_BulletFont.ForRow(row), this.BulletFont.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_Bullet.ForRow(row), this.Bullet.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_BulletFontSize.ForRow(row), this.BulletFontSize.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_LocalizeBulletFont.ForRow(row), this.LocBulletFont.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_TextPosAfterBullet.ForRow(row), this.TextPosAfterBullet.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_Flags.ForRow(row), this.Flags.Formula);
            func(VA.ShapeSheet.SRCConstants.Para_BulletStr.ForRow(row), this.BulletString.Formula);
        }

        public static IList<List<ParagraphCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return _GetCells(page, shapeids, query, query.GetCells);
        }

        public static IList<ParagraphCells> GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return _GetCells(shape, query, query.GetCells);
        }

        private static ParagraphFormatCellQuery _mCellQuery;
        private static ParagraphFormatCellQuery get_query()
        {
            _mCellQuery = _mCellQuery ?? new ParagraphFormatCellQuery();
            return _mCellQuery;
        }

        class ParagraphFormatCellQuery : VA.ShapeSheet.Query.CellQuery
        {
            public VA.ShapeSheet.Query.CellQuery.Column Bullet { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column BulletFont { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column BulletFontSize { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column BulletString { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column Flags { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column HorzAlign { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column IndentFirst { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column IndentLeft { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column IndentRight { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column LocalizeBulletFont { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column SpaceAfter { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column SpaceBefore { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column SpaceLine { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column TextPosAfterBullet { get; set; }

            public ParagraphFormatCellQuery() 
            {
                var sec = this.Sections.Add(IVisio.VisSectionIndices.visSectionParagraph);
                Bullet = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Para_Bullet, "BulletIndex");
                BulletFont = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Para_BulletFont, "BulletFont");
                BulletFontSize = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Para_BulletFontSize, "BulletFontSize");
                BulletString = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Para_BulletStr, "BulletString");
                Flags = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Para_Flags, "Flags");
                HorzAlign = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Para_HorzAlign, "HorzAlign");
                IndentFirst = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Para_IndFirst, "IndentFirst");
                IndentLeft = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Para_IndLeft, "IndentLeft");
                IndentRight = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Para_IndRight, "IndentRight");
                LocalizeBulletFont = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Para_LocalizeBulletFont, "LocalizeBulletFont");
                SpaceAfter = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Para_SpAfter, "SpaceAfter");
                SpaceBefore = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Para_SpBefore, "SpaceBefore");
                SpaceLine = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Para_SpLine, "SpaceLine");
                TextPosAfterBullet = sec.Columns.Add(VA.ShapeSheet.SRCConstants.Para_TextPosAfterBullet, "TextPosAfterBullet");
            }

            public ParagraphCells GetCells(VA.ShapeSheet.CellData<double>[] row)
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
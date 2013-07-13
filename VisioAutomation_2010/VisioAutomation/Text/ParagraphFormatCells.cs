using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Text
{
    public class ParagraphFormatCells : VA.ShapeSheet.CellGroups.CellGroupMultiRowEx
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

        public static IList<List<ParagraphFormatCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return _GetCells(page, shapeids, query, query.GetCells);
        }

        public static IList<ParagraphFormatCells> GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return _GetCells(shape, query, query.GetCells);
        }

        private static ParagraphFormatQuery m_query;
        private static ParagraphFormatQuery get_query()
        {
            m_query = m_query ?? new ParagraphFormatQuery();
            return m_query;
        }

        private static ParagraphFormatCells get_cells_from_row(ParagraphFormatQuery query, VA.ShapeSheet.Data.Table<VA.ShapeSheet.CellData<double>> table, int row)
        {
            var cells = new ParagraphFormatCells();
            cells.IndentFirst = table[row,query.IndentFirst];
            cells.IndentLeft = table[row,query.IndentLeft];
            cells.IndentRight = table[row,query.IndentRight];
            cells.SpacingAfter = table[row,query.SpaceAfter];
            cells.SpacingBefore = table[row,query.SpaceBefore];
            cells.SpacingLine = table[row,query.SpaceLine];
            cells.HorizontalAlign = table[row,query.HorzAlign].ToInt();
            cells.Bullet = table[row,query.Bullet].ToInt();
            cells.BulletFont = table[row,query.BulletFont].ToInt();
            cells.BulletFontSize = table[row,query.BulletFontSize].ToInt();
            cells.LocBulletFont = table[row,query.LocalizeBulletFont].ToInt();
            cells.TextPosAfterBullet= table[row,query.TextPosAfterBullet];
            cells.Flags= table[row,query.Flags].ToInt();
            cells.BulletString = ""; // TODO: Figure out some way of getting this
            return cells;
        }

        class ParagraphFormatQuery : VA.ShapeSheet.Query.QueryEx
        {
            public int Bullet { get; set; }
            public int BulletFont { get; set; }
            public int BulletFontSize { get; set; }
            public int BulletString { get; set; }
            public int Flags { get; set; }
            public int HorzAlign { get; set; }
            public int IndentFirst { get; set; }
            public int IndentLeft { get; set; }
            public int IndentRight { get; set; }
            public int LocalizeBulletFont { get; set; }
            public int SpaceAfter { get; set; }
            public int SpaceBefore { get; set; }
            public int SpaceLine { get; set; }
            public int TextPosAfterBullet { get; set; }

            public ParagraphFormatQuery() 
            {
                var sec = this.AddSection(IVisio.VisSectionIndices.visSectionParagraph);
                Bullet = sec.AddColumn(VA.ShapeSheet.SRCConstants.Para_Bullet, "BulletIndex");
                BulletFont = sec.AddColumn(VA.ShapeSheet.SRCConstants.Para_BulletFont, "BulletFont");
                BulletFontSize = sec.AddColumn(VA.ShapeSheet.SRCConstants.Para_BulletFontSize, "BulletFontSize");
                BulletString = sec.AddColumn(VA.ShapeSheet.SRCConstants.Para_BulletStr, "BulletString");
                Flags = sec.AddColumn(VA.ShapeSheet.SRCConstants.Para_Flags, "Flags");
                HorzAlign = sec.AddColumn(VA.ShapeSheet.SRCConstants.Para_HorzAlign, "HorzAlign");
                IndentFirst = sec.AddColumn(VA.ShapeSheet.SRCConstants.Para_IndFirst, "IndentFirst");
                IndentLeft = sec.AddColumn(VA.ShapeSheet.SRCConstants.Para_IndLeft, "IndentLeft");
                IndentRight = sec.AddColumn(VA.ShapeSheet.SRCConstants.Para_IndRight, "IndentRight");
                LocalizeBulletFont = sec.AddColumn(VA.ShapeSheet.SRCConstants.Para_LocalizeBulletFont, "LocalizeBulletFont");
                SpaceAfter = sec.AddColumn(VA.ShapeSheet.SRCConstants.Para_SpAfter, "SpaceAfter");
                SpaceBefore = sec.AddColumn(VA.ShapeSheet.SRCConstants.Para_SpBefore, "SpaceBefore");
                SpaceLine = sec.AddColumn(VA.ShapeSheet.SRCConstants.Para_SpLine, "SpaceLine");
                TextPosAfterBullet = sec.AddColumn(VA.ShapeSheet.SRCConstants.Para_TextPosAfterBullet, "TextPosAfterBullet");
            }

            public ParagraphFormatCells GetCells(VA.ShapeSheet.CellData<double>[] row)
            {
                var cells = new ParagraphFormatCells();
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
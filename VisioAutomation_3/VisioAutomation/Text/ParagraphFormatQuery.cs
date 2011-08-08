using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Text
{
    class ParagraphFormatQuery : VA.ShapeSheet.Query.SectionQuery
    {
        public VA.ShapeSheet.Query.SectionQueryColumn BulletIndex { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn BulletFont { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn BulletFontSize { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn BulletString { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Flags { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn HorzAlign { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn IndentFirst { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn IndentLeft { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn IndentRight { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn LocalizeBulletFont { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn SpaceAfter { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn SpaceBefore { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn SpaceLine { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn TextPosAfterBullet { get; set; }

        public ParagraphFormatQuery() :
            base(IVisio.VisSectionIndices.visSectionParagraph)
        {
            BulletIndex = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_Bullet, "BulletIndex");
            BulletFont = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_BulletFont, "BulletFont");
            BulletFontSize = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_BulletFontSize, "BulletFontSize");
            BulletString = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_BulletStr, "BulletString");
            Flags = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_Flags, "Flags");
            HorzAlign = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_HorzAlign, "HorzAlign");
            IndentFirst = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_IndFirst, "IndentFirst");
            IndentLeft = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_IndLeft, "IndentLeft");
            IndentRight = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_IndRight, "IndentRight");
            LocalizeBulletFont = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_LocalizeBulletFont, "LocalizeBulletFont");
            SpaceAfter = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_SpAfter, "SpaceAfter");
            SpaceBefore = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_SpBefore, "SpaceBefore");
            SpaceLine = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_SpLine, "SpaceLine");
            TextPosAfterBullet = this.AddColumn(VA.ShapeSheet.SRCConstants.Para_TextPosAfterBullet, "TextPosAfterBullet");
        }
    }
}
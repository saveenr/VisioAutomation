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
            BulletIndex = this.AddColumn(IVisio.VisCellIndices.visBulletIndex, "BulletIndex");
            BulletFont = this.AddColumn(IVisio.VisCellIndices.visBulletFont, "BulletFont");
            BulletFontSize = this.AddColumn(IVisio.VisCellIndices.visBulletFontSize, "BulletFontSize");
            BulletString = this.AddColumn(IVisio.VisCellIndices.visBulletString, "BulletString");
            Flags = this.AddColumn(IVisio.VisCellIndices.visFlags, "Flags");
            HorzAlign = this.AddColumn(IVisio.VisCellIndices.visHorzAlign, "HorzAlign");
            IndentFirst = this.AddColumn(IVisio.VisCellIndices.visIndentFirst, "IndentFirst");
            IndentLeft = this.AddColumn(IVisio.VisCellIndices.visIndentLeft, "IndentLeft");
            IndentRight = this.AddColumn(IVisio.VisCellIndices.visIndentRight, "IndentRight");
            LocalizeBulletFont = this.AddColumn(IVisio.VisCellIndices.visLocalizeBulletFont, "LocalizeBulletFont");
            SpaceAfter = this.AddColumn(IVisio.VisCellIndices.visSpaceAfter, "SpaceAfter");
            SpaceBefore = this.AddColumn(IVisio.VisCellIndices.visSpaceBefore, "SpaceBefore");
            SpaceLine = this.AddColumn(IVisio.VisCellIndices.visSpaceLine, "SpaceLine");
            TextPosAfterBullet = this.AddColumn(IVisio.VisCellIndices.visTextPosAfterBullet, "TextPosAfterBullet");
        }
    }
}
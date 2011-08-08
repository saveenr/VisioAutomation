using VA=VisioAutomation;
using IVisio =Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    class CharacterFormatQuery : VA.ShapeSheet.Query.SectionQuery
    {
        public VA.ShapeSheet.Query.SectionQueryColumn Font { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Style { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Color { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Size { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Trans { get; set; }

        public CharacterFormatQuery() :
            base(IVisio.VisSectionIndices.visSectionCharacter)
        {
            Color = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Color, "Color");
            Trans = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_ColorTrans, "Trans");
            Font = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Font, "Font");
            Size = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Size, "Size");
            Style = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Style, "Style");
        }
    }
}
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
            Color = this.AddColumn(IVisio.VisCellIndices.visCharacterColor, "Color");
            Trans = this.AddColumn(IVisio.VisCellIndices.visCharacterColorTrans, "Trans");
            Font = this.AddColumn(IVisio.VisCellIndices.visCharacterFont, "Font");
            Size = this.AddColumn(IVisio.VisCellIndices.visCharacterSize, "Size");
            Style = this.AddColumn(IVisio.VisCellIndices.visCharacterStyle, "Style");
        }
    }
}
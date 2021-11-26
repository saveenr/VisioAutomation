using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomation.Models.Layouts.Container
{
    public class Formatting
    {
        public Shapes.FormatCells FormatCells;
        public VisioAutomation.Text.CharacterFormatCells CharacterFormatCells;
        public VisioAutomation.Text.ParagraphFormatCells ParagraphFormatCells;
        public VisioAutomation.Text.TextBlockCells TextBlockCells;

        public Formatting()
        {
            this.FormatCells = new Shapes.FormatCells();
            this.CharacterFormatCells = new VisioAutomation.Text.CharacterFormatCells();
            this.ParagraphFormatCells = new VisioAutomation.Text.ParagraphFormatCells();
            this.TextBlockCells = new VisioAutomation.Text.TextBlockCells();
        }

        public void Apply(SidSrcWriter writer, short shapeid_label, short shapeid_box)
        {

            writer.SetValues(shapeid_label, this.ParagraphFormatCells, 0);
            writer.SetValues(shapeid_label, this.CharacterFormatCells, 0);
            writer.SetValues(shapeid_box, this.FormatCells);
            writer.SetValues(shapeid_box, this.TextBlockCells);
        }
    }
}

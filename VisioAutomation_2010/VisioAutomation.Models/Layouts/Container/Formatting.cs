using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomation.Models.Layouts.Container
{
    public class Formatting
    {
        public Shapes.ShapeFormatCells ShapeFormatCells;
        public VisioAutomation.Text.CharacterFormatCells CharacterFormatCells;
        public VisioAutomation.Text.ParagraphFormatCells ParagraphFormatCells;
        public VisioAutomation.Text.TextBlockCells TextBlockCells;

        public Formatting()
        {
            this.ShapeFormatCells = new Shapes.ShapeFormatCells();
            this.CharacterFormatCells = new VisioAutomation.Text.CharacterFormatCells();
            this.ParagraphFormatCells = new VisioAutomation.Text.ParagraphFormatCells();
            this.TextBlockCells = new VisioAutomation.Text.TextBlockCells();
        }

        public void Apply(SidSrcWriter writer, short shapeid_label, short shapeid_box)
        {

            writer.SetFormulas(shapeid_label, this.ParagraphFormatCells, 0);
            writer.SetFormulas(shapeid_label, this.CharacterFormatCells, 0);
            writer.SetFormulas(shapeid_box, this.ShapeFormatCells);
            writer.SetFormulas(shapeid_box, this.TextBlockCells);
        }
    }
}

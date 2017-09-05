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
            this.CharacterFormatCells.SetFormulas(writer, shapeid_label, 0);
            this.ParagraphFormatCells.SetFormulas(writer, shapeid_label, 0);
            this.ShapeFormatCells.SetFormulas(writer, shapeid_box);
            this.TextBlockCells.SetFormulas(writer, shapeid_label);
        }
    }
}

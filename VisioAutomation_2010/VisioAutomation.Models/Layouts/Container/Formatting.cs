using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomation.Models.Layouts.Container
{
    public class Formatting
    {
        public Shapes.ShapeFormatCells ShapeFormatCells;
        public VisioAutomation.Text.CharacterCells CharacterCells;
        public VisioAutomation.Text.ParagraphCells ParagraphCells;
        public VisioAutomation.Text.TextBlockCells TextBlockCells;

        public Formatting()
        {
            this.ShapeFormatCells = new Shapes.ShapeFormatCells();
            this.CharacterCells = new VisioAutomation.Text.CharacterCells();
            this.ParagraphCells = new VisioAutomation.Text.ParagraphCells();
            this.TextBlockCells = new VisioAutomation.Text.TextBlockCells();
        }

        public void Apply(SidSrcWriter writer, short shapeid_label, short shapeid_box)
        {

            writer.SetValues(shapeid_label, this.ParagraphCells, 0);
            writer.SetValues(shapeid_label, this.CharacterCells, 0);
            writer.SetValues(shapeid_box, this.ShapeFormatCells);
            writer.SetValues(shapeid_box, this.TextBlockCells);
        }
    }
}

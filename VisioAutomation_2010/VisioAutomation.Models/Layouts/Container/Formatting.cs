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

        public void Apply(FormulaWriterSIDSRC writer, short shapeid_label, short shapeid_box)
        {
            this.CharacterCells.SetFormulas(shapeid_label, writer, 0);
            this.ParagraphCells.SetFormulas(shapeid_label, writer, 0);
            this.ShapeFormatCells.SetFormulas(shapeid_box, writer);
            this.TextBlockCells.SetFormulas(shapeid_label, writer);
        }
    }
}

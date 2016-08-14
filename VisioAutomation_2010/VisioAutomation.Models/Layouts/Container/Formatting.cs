using VisioAutomation.ShapeSheet.Update;

namespace VisioAutomation.Models.Layouts.Container
{
    public class Formatting
    {
        public Shapes.FormatCells FormatCells;
        public VisioAutomation.Text.CharacterCells CharacterCells;
        public VisioAutomation.Text.ParagraphCells ParagraphCells;
        public VisioAutomation.Text.TextBlockCells TextBlockCells;

        public Formatting()
        {
            this.FormatCells = new Shapes.FormatCells();
            this.CharacterCells = new VisioAutomation.Text.CharacterCells();
            this.ParagraphCells = new VisioAutomation.Text.ParagraphCells();
            this.TextBlockCells = new VisioAutomation.Text.TextBlockCells();
        }

        public void Apply(Update update, short shapeid_label, short shapeid_box)
        {
            update.SetFormulas(shapeid_label, this.CharacterCells, 0);
            update.SetFormulas(shapeid_label, this.ParagraphCells, 0);
            update.SetFormulas(shapeid_box, this.FormatCells);
            update.SetFormulas(shapeid_label, this.TextBlockCells);
        }
    }
}

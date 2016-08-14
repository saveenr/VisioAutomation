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
            this.CharacterCells.SetFormulas(shapeid_label, update, 0);
            this.ParagraphCells.SetFormulas(shapeid_label, update, 0);
            this.FormatCells.SetFormulas(shapeid_box, update);
            this.TextBlockCells.SetFormulas(shapeid_label, update);
        }
    }
}

using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Models.ContainerLayout
{
    public class Formatting
    {
        public VA.Shapes.FormatCells FormatCells;
        public VA.Text.CharacterCells CharacterCells;
        public VA.Text.ParagraphCells ParagraphCells;
        public VA.Text.TextCells TextCells;

        public Formatting()
        {
            this.FormatCells = new VA.Shapes.FormatCells();
            this.CharacterCells = new VA.Text.CharacterCells();
            this.ParagraphCells = new VA.Text.ParagraphCells();
            this.TextCells = new VA.Text.TextCells();
        }

        public void Apply(VA.ShapeSheet.Update update, short shapeid_label, short shapeid_box)
        {
            update.SetFormulas(shapeid_label, this.CharacterCells, 0);
            update.SetFormulas(shapeid_label, this.ParagraphCells, 0);
            update.SetFormulas(shapeid_box, this.FormatCells);
            update.SetFormulas(shapeid_label, this.TextCells);
        }
    }
}

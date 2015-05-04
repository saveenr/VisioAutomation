using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Models.ContainerLayout
{
    public class Formatting
    {
        public Shapes.FormatCells FormatCells;
        public Text.CharacterCells CharacterCells;
        public Text.ParagraphCells ParagraphCells;
        public Text.TextCells TextCells;

        public Formatting()
        {
            this.FormatCells = new Shapes.FormatCells();
            this.CharacterCells = new Text.CharacterCells();
            this.ParagraphCells = new Text.ParagraphCells();
            this.TextCells = new Text.TextCells();
        }

        public void Apply(ShapeSheet.Update update, short shapeid_label, short shapeid_box)
        {
            update.SetFormulas(shapeid_label, this.CharacterCells, 0);
            update.SetFormulas(shapeid_label, this.ParagraphCells, 0);
            update.SetFormulas(shapeid_box, this.FormatCells);
            update.SetFormulas(shapeid_label, this.TextCells);
        }
    }
}

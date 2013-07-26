using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Layout.Models.ContainerLayout
{
    public class Formatting
    {
        public VA.Format.ShapeFormatCells ShapeFormatCells;
        public VA.Text.CharacterFormatCells CharacterFormatCells;
        public VA.Text.ParagraphFormatCells ParagraphFormatCells;
        public VA.Text.TextCells TextCells;

        public Formatting()
        {
            this.ShapeFormatCells = new VA.Format.ShapeFormatCells();
            this.CharacterFormatCells = new VA.Text.CharacterFormatCells();
            this.ParagraphFormatCells = new VA.Text.ParagraphFormatCells();
            this.TextCells = new VA.Text.TextCells();
        }

        public void Apply(VA.ShapeSheet.Update update, short shapeid_label, short shapeid_box)
        {
            update.SetFormulasForRow(shapeid_label, this.CharacterFormatCells, 0);
            update.SetFormulasForRow(shapeid_label, this.ParagraphFormatCells, 0);
            update.SetFormulas(shapeid_box, this.ShapeFormatCells);
            update.SetFormulas(shapeid_label, this.TextCells);
        }
    }
}

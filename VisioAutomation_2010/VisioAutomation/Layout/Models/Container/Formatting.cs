using VisioAutomation.Format;
using VisioAutomation.Text;
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
            this.ShapeFormatCells = new ShapeFormatCells();
            this.CharacterFormatCells = new CharacterFormatCells();
            this.ParagraphFormatCells = new ParagraphFormatCells();
            this.TextCells = new TextCells();
        }

        public void Apply(VA.ShapeSheet.Update update, short shapeid, short shapeid2)
        {
            this.CharacterFormatCells.Apply(update, shapeid, 0);
            this.ParagraphFormatCells.Apply(update, shapeid, 0);
            this.ShapeFormatCells.Apply(update, shapeid2);
            this.TextCells.Apply(update, shapeid);
        }
    }
}

using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Forms
{
    public class TextBlock
    {
        public Drawing.Size Size;
        public string Font = "SegoeUI";
        public VisioAutomation.Text.TextBlockCells TextBlockCells;
        public VisioAutomation.Text.ParagraphCells ParagraphCells;
        public Shapes.FormatCells FormatCells;
        public VisioAutomation.Text.CharacterCells CharacterCells;
        public string Text;
        public IVisio.Shape VisioShape;
        public int VisioShapeID;
        public Drawing.Rectangle Rectangle;

        public TextBlock(Drawing.Size size, string text)
        {
            this.Text = text;
            this.Size = size;
            this.TextBlockCells = new VisioAutomation.Text.TextBlockCells();
            this.ParagraphCells = new VisioAutomation.Text.ParagraphCells();
            this.FormatCells = new Shapes.FormatCells();
            this.CharacterCells = new VisioAutomation.Text.CharacterCells();
        }

        public void ApplyFormus(ShapeSheet.Update update)
        {
            short titleshape_id = this.VisioShape.ID16;
            update.SetFormulas(titleshape_id, this.TextBlockCells);
            update.SetFormulas(titleshape_id, this.ParagraphCells, 0);
            update.SetFormulas(titleshape_id, this.CharacterCells, 0);
            update.SetFormulas(titleshape_id, this.FormatCells);
        }
    }
}
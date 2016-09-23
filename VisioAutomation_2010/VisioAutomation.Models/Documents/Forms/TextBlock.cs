using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Documents.Forms
{
    public class TextBlock
    {
        public Drawing.Size Size;
        public string Font = "SegoeUI";
        public VisioAutomation.Text.TextBlockCells TextBlockCells;
        public VisioAutomation.Text.ParagraphCells ParagraphCells;
        public Shapes.ShapeFormatCells FormatCells;
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
            this.FormatCells = new Shapes.ShapeFormatCells();
            this.CharacterCells = new VisioAutomation.Text.CharacterCells();
        }

        public void ApplyFormus(FormulaWriterSIDSRC writer)
        {
            short titleshape_id = this.VisioShape.ID16;
            this.TextBlockCells.SetFormulas(titleshape_id, writer);
            this.ParagraphCells.SetFormulas(titleshape_id, writer, 0);
            this.CharacterCells.SetFormulas(titleshape_id, writer, 0);
            this.FormatCells.SetFormulas(titleshape_id, writer);
        }
    }
}
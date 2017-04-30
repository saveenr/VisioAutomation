using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Documents.Forms
{
    public class TextBlock
    {
        public Geometry.Size Size;
        public string Font = "SegoeUI";
        public VisioAutomation.Text.TextBlockCells TextBlockCells;
        public VisioAutomation.Text.ParagraphFormatCells ParagraphFormatCells;
        public Shapes.ShapeFormatCells FormatCells;
        public VisioAutomation.Text.CharacterFormatCells CharacterFormatCells;
        public string Text;
        public IVisio.Shape VisioShape;
        public int VisioShapeID;
        public Geometry.Rectangle Rectangle;

        public TextBlock(Geometry.Size size, string text)
        {
            this.Text = text;
            this.Size = size;
            this.TextBlockCells = new VisioAutomation.Text.TextBlockCells();
            this.ParagraphFormatCells = new VisioAutomation.Text.ParagraphFormatCells();
            this.FormatCells = new Shapes.ShapeFormatCells();
            this.CharacterFormatCells = new VisioAutomation.Text.CharacterFormatCells();
        }

        public void ApplyFormus(SidSrcWriter writer)
        {
            short titleshape_id = this.VisioShape.ID16;
            this.TextBlockCells.SetFormulas(titleshape_id, writer);
            this.ParagraphFormatCells.SetFormulas(titleshape_id, writer, 0);
            this.CharacterFormatCells.SetFormulas(titleshape_id, writer, 0);
            this.FormatCells.SetFormulas(titleshape_id, writer);
        }
    }
}
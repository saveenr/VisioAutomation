using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Documents.Forms
{
    public class TextBlock
    {
        public VisioAutomation.Geometry.Size Size;
        public string Font = "SegoeUI";
        public VisioAutomation.Text.TextBlockCells TextBlockCells;
        public VisioAutomation.Text.ParagraphFormatCells ParagraphFormatCells;
        public Shapes.ShapeFormatCells FormatCells;
        public VisioAutomation.Text.CharacterFormatCells CharacterFormatCells;
        public string Text;
        public IVisio.Shape VisioShape;
        public int VisioShapeID;
        public VisioAutomation.Geometry.Rectangle Rectangle;

        public TextBlock(VisioAutomation.Geometry.Size size, string text)
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
            writer.SetFormulas(titleshape_id, this.TextBlockCells);
            writer.SetFormulas(titleshape_id, this.ParagraphFormatCells, 0);
            writer.SetFormulas(titleshape_id, this.CharacterFormatCells, 0);
            writer.SetFormulas(titleshape_id, this.FormatCells);
        }
    }
}
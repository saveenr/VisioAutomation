using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Documents.Forms
{
    public class TextBlock
    {
        public VisioAutomation.Core.Size Size;
        public string Font = "SegoeUI";
        public VisioAutomation.Text.TextBlockCells TextBlockCells;
        public VisioAutomation.Text.ParagraphCells ParagraphCells;
        public Shapes.ShapeFormatCells ShapeFormatCells;
        public VisioAutomation.Text.CharacterCells CharacterCells;
        public string Text;
        public IVisio.Shape VisioShape;
        public int VisioShapeID;
        public VisioAutomation.Core.Rectangle Rectangle;

        public TextBlock(VisioAutomation.Core.Size size, string text)
        {
            this.Text = text;
            this.Size = size;
            this.TextBlockCells = new VisioAutomation.Text.TextBlockCells();
            this.ParagraphCells = new VisioAutomation.Text.ParagraphCells();
            this.ShapeFormatCells = new Shapes.ShapeFormatCells();
            this.CharacterCells = new VisioAutomation.Text.CharacterCells();
        }

        public void ApplyFormus(SidSrcWriter writer)
        {
            short title_shapeid = this.VisioShape.ID16;
            writer.SetValues(title_shapeid, this.TextBlockCells);
            writer.SetValues(title_shapeid, this.ParagraphCells, 0);
            writer.SetValues(title_shapeid, this.CharacterCells, 0);
            writer.SetValues(title_shapeid, this.ShapeFormatCells);
        }
    }
}
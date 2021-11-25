using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Documents.Forms
{
    public class FormPage
    {
        public string Name;
        public VisioAutomation.Core.Size Size;
        public PageMargin PageMargin;
        public IVisio.Page VisioPage;

        public double TitleTextSize { get; set; }
        public double BodyParaSpacingAfter { get; set; }
        public double BodyTextSize { get; set; }
        public string DefaultFont { get; set; }
        public string Title;
        public string Body;

        public FormPage()
        {
            this.Size = new VisioAutomation.Core.Size(8.5, 11);
            this.PageMargin = new PageMargin(0.5, 0.5, 0.5, 0.5);
            this.DefaultFont = "Segoe UI";
            this.BodyTextSize = 8.0;
            this.BodyParaSpacingAfter = 0.0;
            this.TitleTextSize = 15.0;

        }

        internal IVisio.Page Draw(FormRenderingContext context)
        {
            var r = new InteractiveRenderer(context.Document);
            this.VisioPage = r.CreatePage(this);
            context.Page = this.VisioPage;

            var titleblock = new TextBlock(new VisioAutomation.Core.Size(7.5, 0.5), this.Title);

            int fontid = context.GetFontID(this.DefaultFont);
            titleblock.TextBlockCells.VerticalAlign = 0;
            titleblock.ParagraphFormatCells.HorizontalAlign = 0;
            titleblock.FormatCells.LineWeight = 0;
            titleblock.FormatCells.LinePattern = 0;
            titleblock.CharacterFormatCells.Font = fontid;
            titleblock.CharacterFormatCells.Size = get_pt_string(this.TitleTextSize);



            // Draw the shapes
            var titleshape = r.AddShape(titleblock);
            r.Linefeed();

            double body_height = r.GetDistanceToBottomMargin();
            var bodyblock = new TextBlock(new VisioAutomation.Core.Size(7.5, body_height), this.Body);
            bodyblock.ParagraphFormatCells.HorizontalAlign = 0;
            bodyblock.ParagraphFormatCells.SpacingAfter = get_pt_string(this.BodyParaSpacingAfter);
            bodyblock.CharacterFormatCells.Font = fontid;
            bodyblock.CharacterFormatCells.Size = get_pt_string(this.BodyTextSize);
            bodyblock.FormatCells.LineWeight = 0;
            bodyblock.FormatCells.LinePattern = 0;
            bodyblock.TextBlockCells.VerticalAlign = 0;
            bodyblock.FormatCells.LineWeight = 0;
            bodyblock.FormatCells.LinePattern = 0;

            var bodyshape = r.AddShape(bodyblock);
            r.Linefeed();

            r.Finish();
            return this.VisioPage;
        }

        private static string get_pt_string(double size)
        {
            return string.Format("{0}pt", size);
        }
    }
}
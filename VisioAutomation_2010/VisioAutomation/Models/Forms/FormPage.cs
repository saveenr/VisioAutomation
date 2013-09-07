using System.Collections.Generic;
using VisioAutomation.Extensions;
using VisioAutomation.Shapes;
using VisioAutomation.Text;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Forms
{
    public class TextBlock
    {
        public VA.Drawing.Size Size;
        public string Font = "SegoeUI";
        public VA.Text.TextCells Textcells;
        public VA.Text.ParagraphCells ParagraphCells;
        public VA.Shapes.FormatCells FormatCells;
        public VA.Text.CharacterCells CharacterCells;
        public string Text;
        public IVisio.Shape VisioShape;

        public TextBlock(VA.Drawing.Size size,string text)
        {
            this.Text = text;
            this.Size = size;
            this.Textcells = new TextCells();
            this.ParagraphCells = new ParagraphCells();
            this.FormatCells = new FormatCells();
            this.CharacterCells = new CharacterCells();

            this.Textcells.VerticalAlign = 0;
            this.ParagraphCells.HorizontalAlign = 0;

            this.FormatCells.LineWeight = 0;
            this.FormatCells.LinePattern = 0;

        }

        public void ApplyFormus(VA.ShapeSheet.Update update)
        {
            short titleshape_id = this.VisioShape.ID16;
            update.SetFormulas(titleshape_id, this.Textcells);
            update.SetFormulasForRow(titleshape_id, this.ParagraphCells, 0);
            update.SetFormulasForRow(titleshape_id, this.CharacterCells, 0);
            update.SetFormulas(titleshape_id, this.FormatCells);
        }
    }

    public class InteractiveDocumentRenderer
    {
        private IVisio.Pages Pages;
        public VA.Drawing.Point InsertionPoint;

        private double leftmargin;
        private double topmargin;

        private double heightacc ;
        private IVisio.Page page;
        public List<TextBlock> Blocks; 
        public InteractiveDocumentRenderer(IVisio.Document doc)
        {
            this.Pages = doc.Pages;
            this.Blocks = new List<TextBlock>();
        }

        public IVisio.Page AddPage(string name, VA.Drawing.Size size, VA.Pages.PageCells pagecells)
        {
            this.page = this.Pages.Add();
            this.page.Name = name;

            // Update the Page Cells
            var pagesheet = page.PageSheet;
            var pageupdate = new VA.ShapeSheet.Update();
            pagecells.PageWidth = size.Width;
            pagecells.PageHeight = size.Height;
            pageupdate.Execute(pagesheet);

            var rpagecells = VA.Pages.PageCells.GetCells(pagesheet);

            this.leftmargin = rpagecells.PageLeftMargin.Result;
            this.topmargin = size.Height - rpagecells.PageTopMargin.Result;

            this.Reset();
            return this.page;
        }

        public void Reset()
        {
            this.Blocks = new List<TextBlock>();
            this.InsertionPoint = new VA.Drawing.Point(leftmargin, topmargin);
          
        }

        public IVisio.Shape AddShape(TextBlock block)
        {
            var ll = new VA.Drawing.Point(this.InsertionPoint.X, this.InsertionPoint.Y-block.Size.Height);
            var tr  = new VA.Drawing.Point(this.InsertionPoint.X+block.Size.Width, this.InsertionPoint.Y);
            var rect = new VA.Drawing.Rectangle(ll, tr);
            var titleshape = this.page.DrawRectangle(rect);
            block.VisioShape = titleshape;
            if (block.Text != null)
            {
                titleshape.Text = block.Text;                
            }

            this.InsertionPoint = this.InsertionPoint.Add(block.Size.Width, 0);
            this.heightacc = System.Math.Max(this.heightacc, block.Size.Height);


            short titleshape_id = titleshape.ID16;

            this.Blocks.Add(block);

            return titleshape;
        }

        public void Linefeed()
        {
            this.InsertionPoint = new VA.Drawing.Point(leftmargin, this.InsertionPoint.Y - this.heightacc);            
        }

        public void CarriageReturn()
        {
            this.InsertionPoint = new VA.Drawing.Point(leftmargin, this.InsertionPoint.Y);
        }

    }

    public class FormPage
    {
        public string Name;
        public VA.Drawing.Size Size;
        public IVisio.Page VisioPage;

        public double TitleTextSize { get; set; }
        public double BodyParaSpacingAfter { get; set; }
        public double BodyTextSize { get; set; }
        public string DefaultFont { get; set; }
        public string Title;
        public string Body;

        public FormPage()
        {
            this.Size = new VA.Drawing.Size(8.5, 11);

            DefaultFont = "Segoe UI";
            BodyTextSize = 8.0;
            BodyParaSpacingAfter = 0.0;
            TitleTextSize = 15.0;

        }

        public IVisio.Page Draw(FormRenderingContext ctx)
        {

            var r = new InteractiveDocumentRenderer(ctx.Document);

            var page_cells = new VA.Pages.PageCells();
            this.VisioPage = r.AddPage(this.Name, this.Size, page_cells);
            ctx.Page = this.VisioPage;

            var titleblock = new TextBlock(new VA.Drawing.Size(7.5, 1.0), this.Title);

            int _fontid = ctx.GetFontID(this.DefaultFont);
            titleblock.Textcells.VerticalAlign = 0;
            titleblock.ParagraphCells.HorizontalAlign = 0;
            titleblock.FormatCells.LineWeight = 0;
            titleblock.FormatCells.LinePattern = 0;
            titleblock.CharacterCells.Font = _fontid;
            titleblock.CharacterCells.Size = get_pt_string(TitleTextSize);

            var bodyblock = new TextBlock(new VA.Drawing.Size(7.5, 9.0), this.Body);
            bodyblock.ParagraphCells.HorizontalAlign = 0;
            bodyblock.ParagraphCells.SpacingAfter = get_pt_string(BodyParaSpacingAfter);
            bodyblock.CharacterCells.Font = _fontid;
            bodyblock.CharacterCells.Size = get_pt_string(BodyTextSize);
            bodyblock.FormatCells.LineWeight = 0;
            bodyblock.FormatCells.LinePattern = 0;

            // Draw the shapes
            var titleshape = r.AddShape(titleblock);
            r.Linefeed();

            var bodyshape = r.AddShape(bodyblock);
            r.Linefeed();

            var update = new VA.ShapeSheet.Update();
            foreach (var block in r.Blocks)
            {
                block.ApplyFormus(update);
            }
            update.Execute(this.VisioPage);

            return this.VisioPage;
        }

        private string get_pt_string(double size)
        {
            return string.Format("{0}pt", size);
        }
    }
}
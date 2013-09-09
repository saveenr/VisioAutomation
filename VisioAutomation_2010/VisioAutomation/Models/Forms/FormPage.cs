using System.Collections.Generic;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Forms
{
    internal class InteractiveDocumentRenderer
    {
        private IVisio.Pages VisioPages;       
        private double CurrentLineHeight ;
        private IVisio.Page page;
        private FormPage FormPage;

        public List<TextBlock> Blocks;
        public VA.Drawing.Point InsertionPoint;
        
        public InteractiveDocumentRenderer(IVisio.Document doc)
        {
            this.VisioPages = doc.Pages;
            this.Blocks = new List<TextBlock>();
        }

        public IVisio.Page CreatePage(FormPage formpage)
        {
            this.FormPage = formpage;

            this.page = this.VisioPages.Add();
            this.page.Name = formpage.Name;

            // Update the Page Cells
            var pagesheet = page.PageSheet;
            var pageupdate = new VA.ShapeSheet.Update();

            var pagecells = new VA.Pages.PageCells();
            pagecells.PageWidth = formpage.Size.Width;
            pagecells.PageHeight = formpage.Size.Height;
            pagecells.PageLeftMargin = formpage.Margin.Left;
            pagecells.PageRightMargin = formpage.Margin.Right;
            pagecells.PageTopMargin = formpage.Margin.Top;
            pagecells.PageBottomMargin = formpage.Margin.Bottom;
            pageupdate.SetFormulas(pagecells);
            pageupdate.Execute(pagesheet);


            this.Reset();
            return this.page;
        }

        public void Reset()
        {
            this.Blocks = new List<TextBlock>();
            ResetInsertionPoint();
        }

        private void ResetInsertionPoint()
        {
            this.InsertionPoint = new VA.Drawing.Point(this.FormPage.Margin.Left,
                this.FormPage.Size.Height - this.FormPage.Margin.Top);
        }

        public IVisio.Shape AddShape(TextBlock block)
        {
            // Remember this Block 
            this.Blocks.Add(block);

            // Calculate the Correct Full Rectangle
            var ll = new VA.Drawing.Point(this.InsertionPoint.X, this.InsertionPoint.Y-block.Size.Height);
            var tr  = new VA.Drawing.Point(this.InsertionPoint.X+block.Size.Width, this.InsertionPoint.Y);
            var rect = new VA.Drawing.Rectangle(ll, tr);

            // Draw the Shape
            var newshape = this.page.DrawRectangle(rect);
            block.VisioShape = newshape;
            block.VisioShapeID = newshape.ID;
            block.Rectangle = rect;

            // Handle Text If Needed
            if (block.Text != null)
            {
                newshape.Text = block.Text;                
            }

            this.AdjustInsertionPoint(block.Size);

            return newshape;
        }

        private void AdjustInsertionPoint(VA.Drawing.Size size)
        {
            this.InsertionPoint = this.InsertionPoint.Add(size.Width, 0);
            this.CurrentLineHeight = System.Math.Max(this.CurrentLineHeight, size.Height);
        }

        public void Linefeed()
        {
            this.InsertionPoint = new VA.Drawing.Point(this.FormPage.Margin.Left, this.InsertionPoint.Y - this.CurrentLineHeight);            
        }

        public void CarriageReturn()
        {
            this.InsertionPoint = new VA.Drawing.Point(this.FormPage.Margin.Left, this.InsertionPoint.Y);
        }

        public double GetDistanceToBottomMargin()
        {
            double ip_y = this.InsertionPoint.Y - this.CurrentLineHeight;
            double bottom_margin_y = this.FormPage.Margin.Bottom;
            double result = ip_y - bottom_margin_y;
            return result;
        }

        public double GetDistanceToRightMargin()
        {
            double ip_x = this.InsertionPoint.X;
            double right_margin_x = this.FormPage.Size.Width - this.FormPage.Margin.Right;
            double result = ip_x - right_margin_x;
            return result;
        }
    }

    public class FormPage
    {
        public string Name;
        public VA.Drawing.Size Size;
        public VA.Drawing.Margin Margin;
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
            this.Margin = new VA.Drawing.Margin(0.5, 0.5, 0.5, 0.5);
            DefaultFont = "Segoe UI";
            BodyTextSize = 8.0;
            BodyParaSpacingAfter = 0.0;
            TitleTextSize = 15.0;

        }

        internal IVisio.Page Draw(FormRenderingContext ctx)
        {
            var r = new InteractiveDocumentRenderer(ctx.Document);
            var page_cells = new VA.Pages.PageCells();
            this.VisioPage = r.CreatePage(this);
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
using System.Collections.Generic;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Documents.Forms
{
    public class InteractiveRenderer
    {
        private readonly IVisio.Pages _visio_pages;
        private double _current_line_height;
        private IVisio.Page _page;
        private FormPage _form_page;

        public List<TextBlock> Blocks;
        public Drawing.Point InsertionPoint;

        public InteractiveRenderer(IVisio.Document doc)
        {
            this._visio_pages = doc.Pages;
            this.Blocks = new List<TextBlock>();
        }

        public IVisio.Page CreatePage(FormPage formpage)
        {
            this._form_page = formpage;

            this._page = this._visio_pages.Add();
            this._page.Name = formpage.Name;

            // Update the Page Cells
            var pagesheet = this._page.PageSheet;
            var writer = new ShapeSheetWriter();

            var pagecells = new Pages.PageCells();
            pagecells.PageWidth = formpage.Size.Width;
            pagecells.PageHeight = formpage.Size.Height;
            pagecells.PageLeftMargin = formpage.Margin.Left;
            pagecells.PageRightMargin = formpage.Margin.Right;
            pagecells.PageTopMargin = formpage.Margin.Top;
            pagecells.PageBottomMargin = formpage.Margin.Bottom;
            pagecells.SetFormulas(writer);

            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(pagesheet);
            writer.Commit(surface);

            this.Reset();
            return this._page;
        }

        public void Reset()
        {
            this.Blocks = new List<TextBlock>();
            this.ResetInsertionPoint();
        }

        private void ResetInsertionPoint()
        {
            this.InsertionPoint = new Drawing.Point(this._form_page.Margin.Left,
                this._form_page.Size.Height - this._form_page.Margin.Top);
        }

        public TextBlock AddShape(double w, double h, string text)
        {
            var tb = new TextBlock(new Drawing.Size(w, h), text);
            this.AddShape(tb);
            return tb;
        }

        public IVisio.Shape AddShape(TextBlock block)
        {
            // Remember this Block 
            this.Blocks.Add(block);

            // Calculate the Correct Full Rectangle
            var ll = new Drawing.Point(this.InsertionPoint.X, this.InsertionPoint.Y - block.Size.Height);
            var tr = new Drawing.Point(this.InsertionPoint.X + block.Size.Width, this.InsertionPoint.Y);
            var rect = new Drawing.Rectangle(ll, tr);

            // Draw the Shape
            var newshape = this._page.DrawRectangle(rect);
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

        public void Finish()
        {
            var writer = new ShapeSheetWriter();
            foreach (var block in this.Blocks)
            {
                block.FormatCells.SetFormulas((short)block.VisioShapeID,writer);
                block.TextBlockCells.SetFormulas((short)block.VisioShapeID,writer);
                block.ParagraphCells.SetFormulas((short)block.VisioShapeID, writer, 0);
                block.CharacterCells.SetFormulas((short)block.VisioShapeID, writer, 0);
            }

            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(this._page);
            writer.Commit(surface);
        }

        private void AdjustInsertionPoint(Drawing.Size size)
        {
            this.InsertionPoint = this.InsertionPoint.Add(size.Width, 0);
            this._current_line_height = System.Math.Max(this._current_line_height, size.Height);
        }

        public void Linefeed()
        {
            this.InsertionPoint = new Drawing.Point(this._form_page.Margin.Left, this.InsertionPoint.Y - this._current_line_height);
            this._current_line_height = 0;
        }

        public void Linefeed(double advance)
        {
            this.InsertionPoint = new Drawing.Point(this._form_page.Margin.Left, this.InsertionPoint.Y - this._current_line_height - advance);
            this._current_line_height = 0;
        }

        public void MoveRight(double advance)
        {
            this.InsertionPoint = new Drawing.Point(this.InsertionPoint.X + advance, this.InsertionPoint.Y);
            
        }


        public void CarriageReturn()
        {
            this.InsertionPoint = new Drawing.Point(this._form_page.Margin.Left, this.InsertionPoint.Y);
        }

        public double GetDistanceToBottomMargin()
        {
            double ip_y = this.InsertionPoint.Y - this._current_line_height;
            double bottom_margin_y = this._form_page.Margin.Bottom;
            double result = ip_y - bottom_margin_y;
            return result;
        }

        public double GetDistanceToRightMargin()
        {
            double ip_x = this.InsertionPoint.X;
            double right_margin_x = this._form_page.Size.Width - this._form_page.Margin.Right;
            double result = ip_x - right_margin_x;
            return result;
        }
    }
}
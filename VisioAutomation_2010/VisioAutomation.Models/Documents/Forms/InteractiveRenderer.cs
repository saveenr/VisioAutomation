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
        public VisioAutomation.Core.Point InsertionPoint;

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
            var writer = new SrcWriter();

            var page_fmt_cells = new Pages.FormatCells();
            page_fmt_cells.Width = formpage.Size.Width;
            page_fmt_cells.Height = formpage.Size.Height;

            var page_print_cells = new Pages.PrintCells();
            page_print_cells.LeftMargin = formpage.PageMargin.Left;
            page_print_cells.RightMargin = formpage.PageMargin.Right;
            page_print_cells.TopMargin = formpage.PageMargin.Top;
            page_print_cells.BottomMargin = formpage.PageMargin.Bottom;

            writer.SetValues(page_fmt_cells);
            writer.SetValues(page_print_cells);

            writer.Commit(pagesheet, Core.CellValueType.Formula);

            this.Reset();
            return this._page;
        }

        public void Reset()
        {
            this.Blocks = new List<TextBlock>();
            this._reset_insertion_point();
        }

        private void _reset_insertion_point()
        {
            this.InsertionPoint = new VisioAutomation.Core.Point(this._form_page.PageMargin.Left,
                this._form_page.Size.Height - this._form_page.PageMargin.Top);
        }

        public TextBlock AddShape(double w, double h, string text)
        {
            var tb = new TextBlock(new VisioAutomation.Core.Size(w, h), text);
            this.AddShape(tb);
            return tb;
        }

        public IVisio.Shape AddShape(TextBlock block)
        {
            // Remember this Block 
            this.Blocks.Add(block);

            // Calculate the Correct Full Rectangle
            var ll = new VisioAutomation.Core.Point(this.InsertionPoint.X, this.InsertionPoint.Y - block.Size.Height);
            var tr = new VisioAutomation.Core.Point(this.InsertionPoint.X + block.Size.Width, this.InsertionPoint.Y);
            var rect = new VisioAutomation.Core.Rectangle(ll, tr);

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

            this._adjust_insertion_point(block.Size);

            return newshape;
        }

        public void Finish()
        {
            var writer = new SidSrcWriter();
            foreach (var block in this.Blocks)
            {
                writer.SetValues((short)block.VisioShapeID , block.FormatCells);
                writer.SetValues((short)block.VisioShapeID, block.TextBlockCells);

                writer.SetValues((short)block.VisioShapeID, block.ParagraphFormatCells, 0);
                writer.SetValues((short)block.VisioShapeID, block.CharacterFormatCells, 0);
            }

            writer.Commit(this._page, Core.CellValueType.Formula);
        }

        private void _adjust_insertion_point(VisioAutomation.Core.Size size)
        {
            this.InsertionPoint = this.InsertionPoint.Add(size.Width, 0);
            this._current_line_height = System.Math.Max(this._current_line_height, size.Height);
        }

        public void Linefeed()
        {
            this.InsertionPoint = new VisioAutomation.Core.Point(this._form_page.PageMargin.Left, this.InsertionPoint.Y - this._current_line_height);
            this._current_line_height = 0;
        }

        public void Linefeed(double advance)
        {
            this.InsertionPoint = new VisioAutomation.Core.Point(this._form_page.PageMargin.Left, this.InsertionPoint.Y - this._current_line_height - advance);
            this._current_line_height = 0;
        }

        public void MoveRight(double advance)
        {
            this.InsertionPoint = new VisioAutomation.Core.Point(this.InsertionPoint.X + advance, this.InsertionPoint.Y);
            
        }


        public void CarriageReturn()
        {
            this.InsertionPoint = new VisioAutomation.Core.Point(this._form_page.PageMargin.Left, this.InsertionPoint.Y);
        }

        public double GetDistanceToBottomMargin()
        {
            double ip_y = this.InsertionPoint.Y - this._current_line_height;
            double bottom_margin_y = this._form_page.PageMargin.Bottom;
            double result = ip_y - bottom_margin_y;
            return result;
        }

        public double GetDistanceToRightMargin()
        {
            double ip_x = this.InsertionPoint.X;
            double right_margin_x = this._form_page.Size.Width - this._form_page.PageMargin.Right;
            double result = ip_x - right_margin_x;
            return result;
        }
    }
}
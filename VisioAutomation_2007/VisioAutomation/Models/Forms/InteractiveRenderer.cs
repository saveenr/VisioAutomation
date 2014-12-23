using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Forms
{
    public class InteractiveRenderer
    {
        private IVisio.Pages VisioPages;
        private double CurrentLineHeight;
        private IVisio.Page page;
        private FormPage FormPage;

        public List<TextBlock> Blocks;
        public VA.Drawing.Point InsertionPoint;

        public InteractiveRenderer(IVisio.Document doc)
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

        public TextBlock AddShape(double w, double h, string text)
        {
            var tb = new TextBlock(new VA.Drawing.Size(w, h), text);
            this.AddShape(tb);
            return tb;
        }

        public IVisio.Shape AddShape(TextBlock block)
        {
            // Remember this Block 
            this.Blocks.Add(block);

            // Calculate the Correct Full Rectangle
            var ll = new VA.Drawing.Point(this.InsertionPoint.X, this.InsertionPoint.Y - block.Size.Height);
            var tr = new VA.Drawing.Point(this.InsertionPoint.X + block.Size.Width, this.InsertionPoint.Y);
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

        public void Finish()
        {
            var update = new VA.ShapeSheet.Update();
            foreach (var block in this.Blocks)
            {
                update.SetFormulas((short)block.VisioShapeID,block.FormatCells);
                update.SetFormulas((short)block.VisioShapeID,block.Textcells);
                update.SetFormulas((short)block.VisioShapeID,block.ParagraphCells, (short)0);
                update.SetFormulas((short)block.VisioShapeID,block.CharacterCells, (short)0);
            }
            update.Execute(this.page);
        }

        private void AdjustInsertionPoint(VA.Drawing.Size size)
        {
            this.InsertionPoint = this.InsertionPoint.Add(size.Width, 0);
            this.CurrentLineHeight = System.Math.Max(this.CurrentLineHeight, size.Height);
        }

        public void Linefeed()
        {
            this.InsertionPoint = new VA.Drawing.Point(this.FormPage.Margin.Left, this.InsertionPoint.Y - this.CurrentLineHeight);
            this.CurrentLineHeight = 0;
        }

        public void Linefeed(double advance)
        {
            this.InsertionPoint = new VA.Drawing.Point(this.FormPage.Margin.Left, this.InsertionPoint.Y - this.CurrentLineHeight - advance);
            this.CurrentLineHeight = 0;
        }

        public void MoveRight(double advance)
        {
            this.InsertionPoint = new VA.Drawing.Point(this.InsertionPoint.X + advance, this.InsertionPoint.Y);
            
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
}
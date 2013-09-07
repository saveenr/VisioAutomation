using VisioAutomation.Extensions;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Forms
{
    public class FormPage
    {
        public string Name;
        public VA.Drawing.Size Size;
        public MIVisio.Page VisioPage;

        private readonly VA.Drawing.Rectangle _pagerect;
        private readonly VA.Drawing.Rectangle _pageintrect;
        private readonly VA.Drawing.Rectangle _titlerect;
        private readonly VA.Drawing.Rectangle _bodywith_title_rect;
        private int _fontid;
        private VA.Text.TextCells _textblockformat;
        private VA.Text.ParagraphFormatCells _titleParaFmt;
        private VA.Shapes.FormatCells _titleFormat;
        private VA.Text.CharacterFormatCells _titleCharFmt;
        private VA.Text.ParagraphFormatCells _bodyParaFmt;
        private VA.Text.CharacterFormatCells _bodyCharFmt;
        private VA.Shapes.FormatCells _bodyFormat;

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
            _pagerect = new VA.Drawing.Rectangle(new VA.Drawing.Point(0, 0), this.Size);
            _pageintrect = new VA.Drawing.Rectangle(_pagerect.LowerLeft.Add(0.5, 0.5),
                _pagerect.UpperRight.Subtract(0.5, 0.5));

            _titlerect = new VA.Drawing.Rectangle(_pagerect.UpperLeft.Add(0.5, -1.0), _pageintrect.UpperRight);
            _bodywith_title_rect = new VA.Drawing.Rectangle(_pageintrect.LowerLeft, _pagerect.UpperRight.Subtract(0.5, 1.0));

        }

        public IVisio.Page Draw(IVisio.Pages pages)
        {
            var page = pages.Add();

            page.Name = this.Name;

            var pagesheet = page.PageSheet;

            var pageupdate = new VA.ShapeSheet.Update();
            var page_cells = new VA.Pages.PageCells();
            page_cells.PageHeight = this.Size.Height;
            page_cells.PageWidth = this.Size.Width;
            pageupdate.Execute(pagesheet);

            this.VisioPage = page;

            var doc = pages.Document;

            var fonts = doc.Fonts;
            var font = fonts[this.DefaultFont];
            _fontid = font.ID;

            _textblockformat = new VA.Text.TextCells();
            _textblockformat.VerticalAlign = 0;

            _titleParaFmt = new VA.Text.ParagraphFormatCells();
            _titleParaFmt.HorizontalAlign = 0;

            _titleFormat = new VA.Shapes.FormatCells();
            _titleFormat.LineWeight = 0;
            _titleFormat.LinePattern = 0;

            _titleCharFmt = new VA.Text.CharacterFormatCells();
            _titleCharFmt.Font = _fontid;
            _titleCharFmt.Size = get_pt_string(TitleTextSize);

            _bodyParaFmt = new VA.Text.ParagraphFormatCells();
            _bodyParaFmt.HorizontalAlign = 0;
            _bodyParaFmt.SpacingAfter = get_pt_string(BodyParaSpacingAfter);

            _bodyCharFmt = new VA.Text.CharacterFormatCells();
            _bodyCharFmt.Font = _fontid;
            _bodyCharFmt.Size = get_pt_string(BodyTextSize);

            _bodyFormat = new VA.Shapes.FormatCells();
            _bodyFormat.LineWeight = 0;
            _bodyFormat.LinePattern = 0;

            // Draw the shapes
            var titleshape = page.DrawRectangle(_titlerect);
            titleshape.Text = this.Title;

            var bodyshape = page.DrawRectangle(_bodywith_title_rect);
            bodyshape.Text = this.Body;

            var update = new VA.ShapeSheet.Update();

            // Set the ShapeSheet props
            short bodyshape_id = bodyshape.ID16;
            short titleshape_id = titleshape.ID16;
            update.SetFormulas(titleshape_id, _textblockformat);
            update.SetFormulasForRow(titleshape_id, this._titleParaFmt, 0);
            update.SetFormulasForRow(titleshape_id, this._titleCharFmt, 0);
            update.SetFormulas(titleshape_id, this._titleFormat);

            update.SetFormulas(bodyshape_id, _textblockformat);
            update.SetFormulasForRow(bodyshape_id, this._bodyCharFmt, 0);
            update.SetFormulasForRow(bodyshape_id, this._bodyParaFmt, 0);
            update.SetFormulas(bodyshape_id, this._bodyFormat);
            update.Execute(page);


            if (this.Body != null)
            {
                bodyshape.Text = this.Body;
            }

            if (this.Title != null)
            {
                titleshape.Text = this.Title;
            }
            //this.VisioBodyShape = bodyshape;
            //this.VisioTitleShape = titleshape;
            this.VisioPage = page;

            return page;
        }

        private string get_pt_string(double size)
        {
            return string.Format("{0}pt", size);
        }
    }
}
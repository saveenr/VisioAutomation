using VisioAutomation.Drawing;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Models.TextDocument
{
    public class TextDocumentBuilder
    {
        private readonly IVisio.Application _app;

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

        public IVisio.Document VisioDocument { get; private set; }
        public Size PageSize { get; private set; }
        public double TitleTextSize { get; set; }
        public double BodyParaSpacingAfter { get; set; }
        public double BodyTextSize { get; set; }
        public string DefaultFont { get; set; }

        public TextDocumentBuilder(IVisio.Application app, VA.Drawing.Size size)
        {
            DefaultFont = "Segoe UI";
            BodyTextSize = 8.0;
            BodyParaSpacingAfter = 0.0;
            TitleTextSize = 15.0;
            _app = app;
            PageSize = size;
            _pagerect = new VA.Drawing.Rectangle(new VA.Drawing.Point(0, 0), PageSize);
            _pageintrect = new VA.Drawing.Rectangle(_pagerect.LowerLeft.Add(0.5, 0.5),
                                                    _pagerect.UpperRight.Subtract(0.5, 0.5));

            _titlerect = new VA.Drawing.Rectangle(_pagerect.UpperLeft.Add(0.5, -1.0), _pageintrect.UpperRight);
            _bodywith_title_rect = new VA.Drawing.Rectangle(_pageintrect.LowerLeft, _pagerect.UpperRight.Subtract(0.5, 1.0));
        }

        public void Start()
        {
            var docs = _app.Documents;
            VisioDocument = docs.Add("");

            var font = VisioDocument.Fonts[this.DefaultFont];
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
        }

        private string get_pt_string(double size)
        {
            return string.Format("{0}pt",size);
        }

        public void Draw(TextPage textpage)
        {
            var page = VisioDocument.Pages.Add();
            page.NameU = textpage.Name;

            // Update the Page ShapeSheet
            // - to set the size
            var page_cells = new VA.Pages.PageCells();
            page_cells.PageHeight = this.PageSize.Height;
            page_cells.PageWidth = this.PageSize.Width;
            var pageupdate = new VA.ShapeSheet.Update();
            pageupdate.Execute(page);

            // Draw the shapes
            var titleshape = page.DrawRectangle(_titlerect);
            titleshape.Text = textpage.Title;

            var bodyshape = page.DrawRectangle(_bodywith_title_rect);
            bodyshape.Text = textpage.Body;

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

            textpage.VisioBodyShape = bodyshape;
            textpage.VisioTitleShape = titleshape;
            textpage.VisioPage = page;
        }

        public void Finish()
        {
            DeleteFirstPage();

            // set the new first page
            var pages = VisioDocument.Pages;
            var first_page = pages[1];

            var app = VisioDocument.Application;
            var active_window = app.ActiveWindow;
            active_window.Page = first_page;
        }

        private void DeleteFirstPage()
        {
            // Delete the empty first page
            var first_page = VisioDocument.Pages[1];
            first_page.Delete(1);
            first_page = null;
        }
    }

}
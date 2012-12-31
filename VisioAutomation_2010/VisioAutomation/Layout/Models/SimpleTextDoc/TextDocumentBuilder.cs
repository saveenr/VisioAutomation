using VisioAutomation.Drawing;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout.Models.SimpleTextDoc
{
    public class TextDocumentBuilder
    {
        private IVisio.Document _visioDocument;
        private IVisio.Application _app;

        private VA.Drawing.Size _pageSize;
        private VA.Drawing.Rectangle _pagerect;
        private VA.Drawing.Rectangle _pageintrect;
        private VA.Drawing.Rectangle _titlerect;
        private VA.Drawing.Rectangle _bodywith_title_rect;
        private int _fontid;
        private VA.Text.TextCells _textblockformat;
        private VA.Text.ParagraphFormatCells _titleParaFmt;
        private VA.Format.ShapeFormatCells _titleFormat;
        private VA.Text.CharacterFormatCells _titleCharFmt;
        private VA.Text.ParagraphFormatCells _bodyParaFmt;
        private VA.Text.CharacterFormatCells _bodyCharFmt;
        private VA.Format.ShapeFormatCells _bodyFormat;

        private double _titleTextSize = 15.0;
        private double _bodyParaSpacingAfter = 0.0;
        private double _bodyTextSize = 8.0;
        private string _defaultFont = "Segoe UI";

        public TextDocumentBuilder(IVisio.Application app, VA.Drawing.Size size)
        {
            _app = app;
            _pageSize = size;
            _pagerect = new VA.Drawing.Rectangle(new VA.Drawing.Point(0, 0), PageSize);
            _pageintrect = new VA.Drawing.Rectangle(_pagerect.LowerLeft.Add(0.5, 0.5),
                                                    _pagerect.UpperRight.Subtract(0.5, 0.5));

            _titlerect = new VA.Drawing.Rectangle(_pagerect.UpperLeft.Add(0.5, -1.0), _pageintrect.UpperRight);
            _bodywith_title_rect = new VA.Drawing.Rectangle(_pageintrect.LowerLeft, _pagerect.UpperRight.Subtract(0.5, 1.0));
        }

        public IVisio.Document VisioDocument
        {
            get { return _visioDocument; }
        }

        public Size PageSize
        {
            get { return _pageSize; }
        }

        public double TitleTextSize
        {
            get { return _titleTextSize; }
            set { _titleTextSize = value; }
        }

        public double BodyParaSpacingAfter
        {
            get { return _bodyParaSpacingAfter; }
            set { _bodyParaSpacingAfter = value; }
        }

        public double BodyTextSize
        {
            get { return _bodyTextSize; }
            set { _bodyTextSize = value; }
        }

        public string DefaultFont
        {
            get { return _defaultFont; }
            set { _defaultFont = value; }
        }

        public void Start()
        {
            var docs = _app.Documents;
            _visioDocument = docs.Add("");

            var font = _visioDocument.Fonts[this.DefaultFont];
            _fontid = font.ID;

            _textblockformat = new VA.Text.TextCells();
            _textblockformat.VerticalAlign = 0;

            _titleParaFmt = new VA.Text.ParagraphFormatCells();
            _titleParaFmt.HorizontalAlign = 0;

            _titleFormat = new VA.Format.ShapeFormatCells();
            _titleFormat.LineWeight = 0;
            _titleFormat.LinePattern = 0;

            _titleCharFmt = new VA.Text.CharacterFormatCells();
            _titleCharFmt.Font = _fontid;
            _titleCharFmt.Size = get_pt_string(_titleTextSize);

            _bodyParaFmt = new VA.Text.ParagraphFormatCells();
            _bodyParaFmt.HorizontalAlign = 0;
            _bodyParaFmt.SpacingAfter = get_pt_string(_bodyParaSpacingAfter);

            _bodyCharFmt = new VA.Text.CharacterFormatCells();
            _bodyCharFmt.Font = _fontid;
            _bodyCharFmt.Size = get_pt_string(_bodyTextSize);

            _bodyFormat = new VA.Format.ShapeFormatCells();
            _bodyFormat.LineWeight = 0;
            _bodyFormat.LinePattern = 0;
        }

        private string get_pt_string(double size)
        {
            return string.Format("{0}pt",size);
        }

        public void Draw(VA.Layout.Models.SimpleTextDoc.TextPage textpage)
        {
            var page = _visioDocument.Pages.Add();
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
            var pages = _visioDocument.Pages;
            var first_page = pages[1];

            var app = _visioDocument.Application;
            var active_window = app.ActiveWindow;
            active_window.Page = first_page;
        }

        private void DeleteFirstPage()
        {
            // Delete the empty first page
            var first_page = _visioDocument.Pages[1];
            first_page.Delete(1);
            first_page = null;
        }
    }

}
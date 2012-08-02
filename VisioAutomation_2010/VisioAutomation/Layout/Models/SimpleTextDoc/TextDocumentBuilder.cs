using Microsoft.Office.Interop.Visio;
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
        private VA.Drawing.Rectangle _bodywithout_title_rect;
        private int _fontid;
        private VA.Text.TextBlockFormatCells _textblockformat;
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
            _bodywithout_title_rect = _pageintrect;

        }

        public Document VisioDocument
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

            _textblockformat = new VA.Text.TextBlockFormatCells();
            _textblockformat.VerticalAlign = 0;

            _titleParaFmt = new VA.Text.ParagraphFormatCells();
            _titleParaFmt.HorizontalAlign = 0;

            _titleFormat = new VA.Format.ShapeFormatCells();
            _titleFormat.LineWeight = 0;
            _titleFormat.LinePattern = 0;

            _titleCharFmt = new VA.Text.CharacterFormatCells();
            _titleCharFmt.Font = _fontid;
            _titleCharFmt.Size = VA.Convert.PointsToInches(_titleTextSize);

            _bodyParaFmt = new VA.Text.ParagraphFormatCells();
            _bodyParaFmt.HorizontalAlign = 0;
            _bodyParaFmt.SpacingAfter = VA.Convert.PointsToInches(_bodyParaSpacingAfter);

            _bodyCharFmt = new VA.Text.CharacterFormatCells();
            _bodyCharFmt.Font = _fontid;
            _bodyCharFmt.Size = VA.Convert.PointsToInches(_bodyTextSize);

            _bodyFormat = new VA.Format.ShapeFormatCells();
            _bodyFormat.LineWeight = 0;
            _bodyFormat.LinePattern = 0;
            
        }

        public void Draw(VA.Layout.Models.SimpleTextDoc.TextPage textpage)
        {
            var page = _visioDocument.Pages.Add();
            page.NameU = textpage.Name;
            VA.Pages.PageHelper.SetSize(page, this.PageSize);

            // Draw the shapes
            var titleshape = page.DrawRectangle(_titlerect);
            titleshape.Text = textpage.Title;

            var bodyshape = page.DrawRectangle(_bodywith_title_rect);
            bodyshape.Text = textpage.Body;

            var update = new VA.ShapeSheet.Update();

            // Set the ShapeSheet props
            short bodyshape_id = bodyshape.ID16;
            short titleshape_id = titleshape.ID16;
            _textblockformat.Apply(update, titleshape_id);
            this._titleParaFmt.Apply(update, titleshape_id, 0);
            this._titleCharFmt.Apply(update, titleshape_id, 0);
            this._titleFormat.Apply(update, titleshape_id);

            _textblockformat.Apply(update, bodyshape_id);
            this._bodyCharFmt.Apply(update, bodyshape_id, 0);
            this._bodyParaFmt.Apply(update, bodyshape_id, 0);
            this._bodyFormat.Apply(update, bodyshape_id);
            update.Execute(page);

            textpage.VisioBodyShape = bodyshape;
            textpage.VisioTitleShape = titleshape;
            textpage.VisioPage = page;
        }

        public void Finish()
        {
            DeleteFirstPage();

            // set the new first page
            var first_page = _visioDocument.Pages[1];
            first_page.Activate();
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
using Microsoft.Office.Interop.Visio;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Experimental.SimpleTextDoc
{
    public class TextPage
    {
        public string Title;
        public string Body;
        public string Name;
    }

    public class TextDocumentBuilder
    {
        private IVisio.Document _visioDocument;

        private VA.Drawing.Size _pagesize;
        private VA.Drawing.Rectangle _pagerect;
        private VA.Drawing.Rectangle _titlerect;
        private VA.Drawing.Rectangle _bodyrect;
        private int _fontid;
        private VA.Text.TextBlockFormatCells _textblockformat;
        private VA.Text.ParagraphFormatCells _titleParaFmt;
        private VA.Format.ShapeFormatCells _titleFormat;
        private VA.Text.CharacterFormatCells _titleCharFmt;
        private VA.Text.ParagraphFormatCells _bodyParaFmt;
        private VA.Text.CharacterFormatCells _bodyCharFmt;
        private VA.Format.ShapeFormatCells _bodyFormat;

        public TextDocumentBuilder(IVisio.Application app)
        {
            _pagesize = new VA.Drawing.Size(8.5, 11);
            _pagerect = new VA.Drawing.Rectangle(new VA.Drawing.Point(0, 0), _pagesize);
            _titlerect = new VA.Drawing.Rectangle(_pagerect.UpperLeft.Add(0.5, -1.0), _pagerect.UpperRight.Subtract(0.5, 0.5));
            _bodyrect = new VA.Drawing.Rectangle(_pagerect.LowerLeft.Add(0.5, 0.5), _pagerect.UpperRight.Subtract(0.5, 1.0));

            var docs = app.Documents;
            _visioDocument = docs.Add("");

            var font = _visioDocument.Fonts["Segoe UI"];
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
            _titleCharFmt.Size = VA.Convert.PointsToInches(15.0);

            _bodyParaFmt = new VA.Text.ParagraphFormatCells();
            _bodyParaFmt.HorizontalAlign = 0;
            _bodyParaFmt.SpacingAfter = VA.Convert.PointsToInches(6.0);

            _bodyCharFmt = new VA.Text.CharacterFormatCells();
            _bodyCharFmt.Font = _fontid;
            _bodyCharFmt.Size = VA.Convert.PointsToInches(8.0);

            _bodyFormat = new VA.Format.ShapeFormatCells();
            _bodyFormat.LineWeight = 0;
            _bodyFormat.LinePattern = 0;
        }

        public Document VisioDocument
        {
            get { return _visioDocument; }
        }

        public void Draw(VA.Experimental.SimpleTextDoc.TextPage xpage)
        {
            var page = _visioDocument.Pages.Add();
            page.NameU = xpage.Name;
            VA.PageHelper.SetSize(page, this._pagesize);

            // Draw the shapes
            var titleshape = page.DrawRectangle(_titlerect);
            titleshape.Text = xpage.Title;

            var bodyshape = page.DrawRectangle(_bodyrect);
            bodyshape.Text = xpage.Body;

            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

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
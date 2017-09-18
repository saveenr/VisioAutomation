using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class TextCommands : CommandSet
    {
        internal TextCommands(Client client) :
            base(client)
        {

        }

        public void Set(VisioScripting.Models.TargetShapes targets, IList<string> texts)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            if (texts == null || texts.Count < 1)
            {
                return;
            }

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Application.NewUndoScope("Set Shape Text"))
            {
                // Apply text to each shape
                // if there are fewer texts than shapes then
                // start reusing the texts from the beginning

                int count = 0;
                foreach (var shape in targets.Shapes)
                {
                    string text = texts[count%texts.Count];
                    if (text != null)
                    {
                        shape.Text = text;
                    }
                    count++;
                }
            }
        }

        public List<string> Get(VisioScripting.Models.TargetShapes targets)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return new List<string>(0);
            }

            var texts = targets.Shapes.Select(s => s.Text).ToList();
            return texts;
        }

        public void SetFont(VisioScripting.Models.TargetShapes targets, string fontname)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return;
            }

            var application = this._client.Application.Get();
            var active_document = application.ActiveDocument;
            var active_doc_fonts = active_document.Fonts;
            var font = active_doc_fonts[fontname];
            var page = this._client.Page.Get();

            var cells = new VisioAutomation.Text. CharacterFormatCells();
            cells.Font = font.ID;

            this._client.ShapeSheet.__SetCells(targets, cells, page);
        }

        public List<VisioAutomation.Text.TextFormat> GetFormat(VisioScripting.Models.TargetShapes targets)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return new List<VisioAutomation.Text.TextFormat>(0);
            }

            var selection = this._client.Selection.Get();
            var shapeids = selection.GetIDs();
            var application = this._client.Application.Get();
            var formats = VisioAutomation.Text.TextFormat.GetFormat(application.ActivePage, shapeids, CellValueType.Formula);
            return formats;
        }

    }
}



using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet;

namespace VisioScripting.Commands
{
    public class TextCommands : CommandSet
    {
        internal TextCommands(Client client) :
            base(client)
        {

        }

        public void SetShapeText(Models.TargetShapes targets, IList<string> texts)
        {
            if (texts == null || texts.Count < 1)
            {
                return;
            }

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(SetShapeText)))
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

        public List<string> GetShapeText(Models.TargetShapes targets)
        {
            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return new List<string>(0);
            }

            var texts = targets.Shapes.Select(s => s.Text).ToList();
            return texts;
        }

        public List<VisioAutomation.Text.TextFormat> GetShapeTextFormat(Models.TargetShapes targets)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return new List<VisioAutomation.Text.TextFormat>(0);
            }

            var pairs = targets.ToShapeIdPairs();
            var application = cmdtarget.Application;
            var formats = VisioAutomation.Text.TextFormat.GetFormat(application.ActivePage, pairs, CellValueType.Formula);
            return formats;
        }
    }
}

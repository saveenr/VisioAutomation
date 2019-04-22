using System.Collections.Generic;
using System.Linq;
using VASS=VisioAutomation.ShapeSheet;

namespace VisioScripting.Commands
{
    public class TextCommands : CommandSet
    {
        internal TextCommands(Client client) :
            base(client)
        {

        }

        public void SetShapeText(TargetShapes targetshapes, IList<string> texts)
        {
            if (texts == null || texts.Count < 1)
            {
                return;
            }

            targetshapes = targetshapes.Resolve(this._client);

            if (targetshapes.Shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(SetShapeText)))
            {
                // Apply text to each shape
                // if there are fewer texts than shapes then
                // start reusing the texts from the beginning

                int count = 0;
                foreach (var shape in targetshapes.Shapes)
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

        public List<string> GetShapeText(TargetShapes targetshapes)
        {
            targetshapes = targetshapes.Resolve(this._client);

            if (targetshapes.Shapes.Count < 1)
            {
                return new List<string>(0);
            }

            var texts = targetshapes.Shapes.Select(s => s.Text).ToList();
            return texts;
        }

        public List<VisioAutomation.Text.TextFormat> GetShapeTextFormat(TargetShapes targetshapes)
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.RequireDocument);
            targetshapes = targetshapes.Resolve(this._client);

            if (targetshapes.Shapes.Count < 1)
            {
                return new List<VisioAutomation.Text.TextFormat>(0);
            }

            var shapeidpairs = targetshapes.ToShapeIDPairs();
            var application = cmdtarget.Application;
            var formats = VisioAutomation.Text.TextFormat.GetFormat(application.ActivePage, shapeidpairs, VASS.CellValueType.Formula);
            return formats;
        }
    }
}

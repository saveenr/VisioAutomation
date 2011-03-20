using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public void SetText( string t)
        {
            this.ScriptingSession.Text.SetText(new string[] { t });
        }

        public void SetText(IEnumerable<string> items)
        {
            this.ScriptingSession.Text.SetText(items);
        }

        public IList<string> GetText()
        {
            return this.ScriptingSession.Text.GetText();
        }

        public IList<string> TextMarkup
        {
            set
            {
                const bool preserve_whitespace = false;
                SetMarkup(value, preserve_whitespace);
            }
        }

        private void SetMarkup(IList<string> value, bool preserve_whitespace)
        {
            var scriptingsession = this.ScriptingSession;
            if (!scriptingsession.HasSelectedShapes())
            {
                return;
            }

            var markup_doms = value
                .Select(s => VA.Text.Markup.TextElement.FromXml(s, preserve_whitespace))
                .ToList();
            var shapes = scriptingsession.Selection.EnumSelectedShapes().ToList();

            using (var undo = scriptingsession.Application.CreateUndoScope())
            {
                for (int i = 0; i < shapes.Count; i++)
                {
                    var shape = shapes[i];
                    var markup_dom = markup_doms[i%markup_doms.Count];
                    markup_dom.SetShapeText(shape);
                }
            }
        }

        private void SetMarkup(VA.Text.Markup.TextElement el, bool preserve_whitespace)
        {
            var scriptingsession = this.ScriptingSession;
            if (!scriptingsession.HasSelectedShapes())
            {
                return;
            }

            var selection = scriptingsession.Selection.GetSelection();
            var shapes = selection.AsEnumerable().ToList();

            using (var undo = scriptingsession.Application.CreateUndoScope())
            {
                foreach (var shape in shapes)
                {
                    el.SetShapeText(shape);
                }
            }
        }

        public bool TextWrapping
        {
            set { this.ScriptingSession.Text.SetTextWrapping(value); }
        }

        public void FitShapeToText()
        {
            this.ScriptingSession.Text.FitShapeToText();
        }

        public void ToogleCase()
        {
            this.ScriptingSession.Text.ToogleCase();
        }

        public void InsertField(VA.Text.Markup.Field field, int begin, int end)
        {
            this.ScriptingSession.Text.InsertField(field, begin, end);
        }

        public void InsertCustomField(int begin, int end, string formula)
        {
            InsertCustomField(begin, end, formula, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
        }

        public void InsertCustomField(int begin, int end, string formula, IVisio.VisFieldFormats format)
        {
            this.ScriptingSession.Text.InsertCustomField(begin, end, formula,
                                               IVisio.VisFieldFormats.visFmtNumGenNoUnits);
        }

        public void ClearCharacterFormat()
        {
            this.ScriptingSession.Text.ClearCharacterFormat();
        }
    }
}
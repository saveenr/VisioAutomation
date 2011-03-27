using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting
{
    public class FormatPainter
    {
        private VA.Format.FormatPaintCache cache = new VA.Format.FormatPaintCache();

        public void Copy(Session scriptingsession)
        {
            var allflags = this.cache.GetAllFormatPaintFlags();
            this.Copy(scriptingsession, allflags);
        }

        public void Copy(Session scriptingsession, VA.Format.FormatCategory category)
        {
            if (!scriptingsession.HasSelectedShapes())
            {
                return;
            }

            var selection = scriptingsession.Selection.GetSelection();
            var shape = selection[1];
            this.cache.CopyFormat(shape, category);
        }

        public void Clear()
        {
            this.cache.Clear();
        }

        public void Paste(Session scriptingsession)
        {
            var allflags = this.cache.GetAllFormatPaintFlags();

            this.Paste(scriptingsession, allflags);
        }

        public void Paste(Session scriptingsession, VA.Format.FormatCategory category)
        {
            if (!scriptingsession.HasSelectedShapes())
            {
                return;
            }

            var selection = scriptingsession.Selection.GetSelection();
            var shapeids = selection.GetIDs();
            var application = scriptingsession.VisioApplication;
            var active_page = application.ActivePage;
            
            this.cache.PasteFormat(active_page, shapeids, category);
        }

    }
}
using System.Collections.Generic;
using System.Xml.Linq;
using VisioAutomation.Extensions;
using System.Linq;
using VA=VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class FormatCommands : SessionCommands
    {
        public FormatCommands(Session session) :
            base(session)
        {

        }

        public void SetFormat(VA.Format.ShapeFormatCells format)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }


            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            var shapes = this.Session.Selection.EnumSelectedShapes().ToList();
            var shapeids = shapes.Select(s => s.ID).ToList();

            foreach (int shapeid in shapeids)
            {
                format.Apply(update, (short) shapeid);
            }

            update.Execute(this.Session.VisioApplication.ActivePage);            
        }

        public void Duplicate(int n)
        {
            if (n < 1)
            {
                throw new System.ArgumentOutOfRangeException("n");
            }
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            // TODO: Add ability to duplicate all the selected shapes, not just the first one
            // this dupicates exactly 1 shape N - times what it
            // it should do is duplicate all M selected shapes N times so that M*N shapes are created

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                var active_window = application.ActiveWindow;
                var selection = active_window.Selection;
                var active_page = application.ActivePage;
                DrawCommandsUtil.CreateDuplicates(active_page, selection[1], n);
            }
        }

        [System.Flags]
        public enum SizeFlags
        {
            Width = 0x1,
            Height = 0x02
        }

        private double? cached_size_width;
        private double? cached_size_height;

        /// <summary>
        /// Caches the resize (the results, not formulas) of a the first currently selected shape
        /// </summary>
        public void CopySize()
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var application = this.Session.VisioApplication;
            var active_window = application.ActiveWindow;
            var selection = active_window.Selection;
            var shape = selection[1];

            var query = new VA.ShapeSheet.Query.CellQuery();
            var width_col = query.AddColumn(VA.ShapeSheet.SRCConstants.Width);
            var height_col = query.AddColumn(VA.ShapeSheet.SRCConstants.Height);
            var queryresults = query.GetResults<double>(shape);

            cached_size_width = queryresults[0, width_col];
            cached_size_height = queryresults[0, height_col];
        }



        /// <summary>
        /// Applies the cached size to the currently selected shapes. If no shapes are selected then nothing happens.
        /// If no size was cached then nothing happens.
        /// </summary>
        /// <param name="flags">Controls if either or both the width and height values are applied during the paste</param>
        public void PasteSize(SizeFlags flags)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            if ((!cached_size_width.HasValue) && (!cached_size_height.HasValue))
            {
                return;
            }

            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            var shapes = this.Session.Selection.EnumSelectedShapes().ToList();
            var shapeids = shapes.Select(s => s.ID).ToList();

            foreach (var shapeid in shapeids)
            {
                if ((flags & SizeFlags.Width) > 0)
                {
                    update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.Width, cached_size_width.Value);
                }

                if ((flags & SizeFlags.Height) > 0)
                {
                    update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.Height, cached_size_height.Value);
                }
            }

            var application = this.Session.VisioApplication;
            var active_page = application.ActivePage;
            update.Execute(active_page);
        }


        private VA.Format.FormatPaintCache cache = new VA.Format.FormatPaintCache();

        public void CopyFormat()
        {
            var allflags = this.cache.GetAllFormatPaintFlags();
            this.CopyFormat(allflags);
        }

        public void CopyFormat(VA.Format.FormatCategory category)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var selection = this.Session.Selection.GetSelection();
            var shape = selection[1];
            this.cache.CopyFormat(shape, category);
        }

        public void ClearFormatCache()
        {
            this.cache.Clear();
        }

        public void PasteFormat()
        {
            var allflags = this.cache.GetAllFormatPaintFlags();

            this.PasteFormat(allflags);
        }

        public void PasteFormat(VA.Format.FormatCategory category)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var selection = this.Session.Selection.GetSelection();
            var shapeids = selection.GetIDs();
            var application = this.Session.VisioApplication;
            var active_page = application.ActivePage;

            this.cache.PasteFormat(active_page, shapeids, category);
        }
    }
}
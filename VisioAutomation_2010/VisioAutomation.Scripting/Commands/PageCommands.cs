using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VisioAutomation.Scripting.Exceptions;
using VisioAutomation.Scripting.View;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Writers;
using VisioAutomation.Utilities;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Commands
{
    public class PageCommands : CommandSet
    {
        internal PageCommands(Client client) :
            base(client)
        {

        }

        public void Set(IVisio.Page page)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var app = this._client.Application.Get();
            this._client.WriteVerbose("Setting Active Page to \"{0}\"", page.Name);
            var window = app.ActiveWindow;
            window.Page = page;
        }

        public void Set(string name)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            var doc = application.ActiveDocument;
            this._client.WriteVerbose("Retrieving Page \"{0}\"", name);
            var pages = doc.Pages;
            var page = pages[name];
            this.Set(page);
        }

        public void Set(int pagenumber)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            var doc = application.ActiveDocument;
            this._client.WriteVerbose("Retrieving Page Number \"{0}\"", pagenumber);
            var pages = doc.Pages;
            var page = pages[pagenumber];
            this.Set(page);
        }
        
        public IVisio.Page Get()
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            return application.ActivePage;
        }

        public void Delete(IList<IVisio.Page> pages, bool renumber)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            if (pages == null)
            {
                throw new System.ArgumentNullException(nameof(pages));
            }

            foreach (var page in pages)
            {
                page.Delete(renumber ? (short) 1 : (short) 0);
            }
        }

        public void Delete(IList<string> names, bool renumber)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            if (names == null)
            {
                throw new System.ArgumentNullException(nameof(names));
            }

            foreach (var name in names)
            {
                var app = this._client.Application.Get();
                var doc = app.ActiveDocument;
                var pages = doc.Pages;

                this._client.WriteVerbose("Retrieving Page for name \"{0}\"",name);
                var page = pages.ItemU[name];
                page.Delete(renumber ? (short)1 : (short)0);
            }
        }

        public Drawing.Size GetSize()
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            var active_page = application.ActivePage;


            var query = new VisioAutomation.ShapeSheet.Queries.Query();
            var col_height = query.AddCell(VisioAutomation.ShapeSheet.SRCConstants.PageHeight, "PageHeight");
            var col_width = query.AddCell(VisioAutomation.ShapeSheet.SRCConstants.PageWidth, "PageWidth");
            var page_surface = new ShapeSheetSurface(active_page.PageSheet);

            var results = query.GetResults<double>(page_surface);
            double height = results.Cells[col_height];
            double width = results.Cells[col_width];
            var s = new Drawing.Size(width, height);
            return s;
        }

        public void SetName(string name)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            if (name.Length < 1)
            {
                throw new System.ArgumentException("name cannot be empty",nameof(name));
            }

            var page = this.Get();
            page.NameU = name;
        }

        public IVisio.Page New(Drawing.Size? size, bool isbackgroundpage)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            var active_document = application.ActiveDocument;
            var pages = active_document.Pages;
            IVisio.Page page = pages.Add();

            using (var undoscope = this._client.Application.NewUndoScope("New Page"))
            {
                if (size.HasValue)
                {
                    this._client.WriteVerbose("Setting page size to {0}", size.Value);
                    this.SetSize(size.Value);
                }

                if (isbackgroundpage)
                {
                    page.Background = 1;
                }
            }

            return page;
        }

        public void SetBackgroundPage(string background_page_name)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            if (background_page_name == null)
            {
                throw new System.ArgumentNullException(nameof(background_page_name));
            }

            var app = this._client.Application.Get();
            var application = app;
            var active_document = application.ActiveDocument;
            var pages = active_document.Pages;
            var names = new HashSet<string>(pages.GetNamesU());
            if (!names.Contains(background_page_name))
            {
                string msg = string.Format("Could not find page with name \"{0}\"", background_page_name);
                throw new VisioOperationException(msg);
            }

            var bgpage = pages.ItemU[background_page_name];
            var fgpage = application.ActivePage;

            // Set the background page
            // Check that the intended background is indeed a background page
            if (bgpage.Background == 0)
            {
                string msg = string.Format("Page \"{0}\" is not a background page", bgpage.Name);
                throw new VisioOperationException(msg);
            }

            // don't allow the page to be set as a background to itself
            if (fgpage == bgpage)
            {
                string msg = "Cannot set page as its own background page";
                throw new VisioOperationException(msg);
            }

            using (var undoscope = this._client.Application.NewUndoScope("Set Background Page"))
            {
                fgpage.BackPage = bgpage;
            }
        }

        public IVisio.Page Duplicate()
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Duplicate Page"))
            {
                var doc = application.ActiveDocument;
                var pages = doc.Pages;
                var src_page = application.ActivePage;
                var new_page = pages.Add();

                var win = application.ActiveWindow;
                win.Page = src_page;
                Pages.PageHelper.Duplicate(src_page, new_page);
                win.Page = new_page;
                return new_page;
            }
        }

        public IVisio.Page Duplicate(IVisio.Document dest_doc)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            if (dest_doc==null)
            {
                throw new System.ArgumentNullException(nameof(dest_doc));
            }

            var application = this._client.Application.Get();

            if (application.ActiveDocument == dest_doc)
            {
                throw new VisioOperationException("dest doc is same as active doc");
            }

            var src_page = application.ActivePage;
            var dest_pages = dest_doc.Pages;
            var dest_page = dest_pages[1];
            Pages.PageHelper.Duplicate(src_page, dest_page);

            // the active window will be to the new document
            var active_window = application.ActiveWindow;
            //active_window.Page = dest_page;

            return dest_page;
        }

        public Pages.PrintPageOrientation GetOrientation()
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            var active_page = application.ActivePage;
            return PageCommands.GetOrientation(active_page);
        }

        private static Pages.PrintPageOrientation GetOrientation(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            var page_sheet = page.PageSheet;
            var src = VisioAutomation.ShapeSheet.SRCConstants.PrintPageOrientation;
            var orientationcell = page_sheet.CellsSRC[src.Section, src.Row, src.Cell];
            int value = orientationcell.ResultInt[IVisio.VisUnitCodes.visNumber, 0];
            return (Pages.PrintPageOrientation)value;
        }

        public void SetOrientation(Pages.PrintPageOrientation orientation)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var app = this._client.Application.Get();
            var application = app;

            var active_page = application.ActivePage;

            if (orientation != Pages.PrintPageOrientation.Landscape && orientation != Pages.PrintPageOrientation.Portrait)
            {
                throw new System.ArgumentOutOfRangeException(nameof(orientation), "must be either Portrait or Landscape");
            }

            var old_orientation = PageCommands.GetOrientation(active_page);

            if (old_orientation == orientation)
            {
                // don't need to do anything
                return;
            }

            var old_size = this.GetSize();

            double new_height = old_size.Width;
            double new_width = old_size.Height;

            var writer = new FormulaWriter();
            writer.SetFormula(VisioAutomation.ShapeSheet.SRCConstants.PageWidth, new_width);
            writer.SetFormula(VisioAutomation.ShapeSheet.SRCConstants.PageHeight, new_height);
            writer.SetFormula(VisioAutomation.ShapeSheet.SRCConstants.PrintPageOrientation, (int)orientation);

            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(active_page.PageSheet);

            using (var undoscope = this._client.Application.NewUndoScope("Set Page Orientation"))
            {
                writer.Commit(surface);
            }
        }



        public void ResizeToFitContents(Drawing.Size bordersize, bool zoom_to_page)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Resize Page to Fit Contents"))
            {
                var active_page = application.ActivePage;
                active_page.ResizeToFitContents(bordersize);
                if (zoom_to_page)
                {
                    this._client.View.Zoom(Zoom.ToPage);
                }
            }
        }

        public void ResetOrigin(IVisio.Page page)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            if (page == null)
            {
                page = application.ActivePage;
            }

            var writer = new FormulaWriter();

            writer.SetFormula(VisioAutomation.ShapeSheet.SRCConstants.XGridOrigin, "0.0");
            writer.SetFormula(VisioAutomation.ShapeSheet.SRCConstants.YGridOrigin, "0.0");
            writer.SetFormula(VisioAutomation.ShapeSheet.SRCConstants.XRulerOrigin, "0.0");
            writer.SetFormula(VisioAutomation.ShapeSheet.SRCConstants.YRulerOrigin, "0.0");

            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(page.PageSheet);

            using (var undoscope = this._client.Application.NewUndoScope("Reset Page Origin"))
            {
                writer.Commit(surface);
            }
        }

        public void SetSize(Drawing.Size new_size)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Set Page Size"))
            {
                var active_page = application.ActivePage;
                var page_sheet = active_page.PageSheet;
                var writer = new FormulaWriter();
                writer.SetFormula(VisioAutomation.ShapeSheet.SRCConstants.PageWidth, new_size.Width);
                writer.SetFormula(VisioAutomation.ShapeSheet.SRCConstants.PageHeight, new_size.Height);

                var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(page_sheet);
                writer.Commit(surface);
            }
        }

        public void SetSize(double? width, double? height)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            if (!width.HasValue && !height.HasValue)
            {
                // nothing to do
                return;
            }

            var old_size = this.GetSize();
            var w = width.GetValueOrDefault(old_size.Width);
            var h = height.GetValueOrDefault(old_size.Height);
            var new_size = new Drawing.Size(w, h);
            this.SetSize(new_size);
        }

        public void SetHeight(double height)
        {
            this.SetSize(null, height);
        }

        public void SetWidth(double width)
        {
            this.SetSize(width, null);
        }

        public void GoTo(PageDirection flags)
        {
            this._client.Application.AssertApplicationAvailable();

            var application = this._client.Application.Get();
            var active_document = application.ActiveDocument;
            var docpages = active_document.Pages;
            if (docpages.Count < 2)
            {
                return;
            }

            var pages = docpages;
            this._GoTo(pages, flags);
        }

        private void _GoTo(IVisio.Pages pages, PageDirection flags)
        {
            this._client.Application.AssertApplicationAvailable();

            if (pages == null)
            {
                throw new System.ArgumentNullException(nameof(pages));
            }

            var app = pages.Application;
            var active_document = app.ActiveDocument;
            if (pages.Document != active_document)
            {
                throw new VisioOperationException("Page.Document is not application's ActiveDocument");
            }

            if (pages.Count < 2)
            {
                throw new VisioOperationException("Only 1 page available. Navigation not possible.");
            }

            var activepage = app.ActivePage;

            int cur_index = activepage.Index;
            const int min_index = 1;
            int max_index = pages.Count;
            int new_index = PageCommands.move_in_range(cur_index, min_index, max_index, flags);
            if (cur_index != new_index)
            {
                var doc_pages = active_document.Pages;
                var page = doc_pages[new_index];

                var active_window = app.ActiveWindow;
                active_window.Page = page;
            }
        }

        internal static int move_in_range(int cur, int min, int max, PageDirection direction)
        {
            if (max < min)
            {
                throw new System.ArgumentOutOfRangeException(nameof(max));
            }

            if (cur < min)
            {
                throw new System.ArgumentOutOfRangeException(nameof(cur));
            }

            if (cur > max)
            {
                throw new System.ArgumentOutOfRangeException(nameof(cur));
            }

            if (direction == PageDirection.Next)
            {
                return System.Math.Min(cur + 1, max);
            }
            else if (direction == PageDirection.Previous)
            {
                return System.Math.Max(cur - 1, min);
            }
            else if (direction == PageDirection.First)
            {
                return min;
            }
            else if (direction == PageDirection.Last)
            {
                return max;
            }
            else
            {
                throw new System.ArgumentOutOfRangeException(nameof(direction));
            }
        }

        public List<IVisio.Shape> GetShapesByID(int[] shapeids)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var page = this._client.Page.Get();
            var shapes = page.Shapes;
            var shapes_list = new List<IVisio.Shape>(shapeids.Length);
            foreach (int id in shapeids)
            {
                var shape = shapes.ItemFromID[id];
                shapes_list.Add(shape);
            }
            return shapes_list;
        }

        public List<IVisio.Shape> GetShapesByName(string[] shapenames)
        {
            return this.GetShapesByName(shapenames, false);
        }

        public List<IVisio.Shape> GetShapesByName(string[] shapenames, bool ignore_bad_names)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var page = this._client.Page.Get();
            var shapes = page.Shapes;

            var cached_shapes_list = new List<IVisio.Shape>(shapes.Count);
            cached_shapes_list.AddRange(shapes.ToEnumerable());
            
            if (shapenames.Contains("*"))
            {
                // if any of the shape names contains a simple wildcard then return all the shapes
                return cached_shapes_list;
            }

            // otherwise we start checking for each name
            var shapes_list = WildcardHelper.FilterObjectsByNames(cached_shapes_list, shapenames, s => s.Name, true, WildcardHelper.FilterAction.Include).ToList();

            return shapes_list;
        }

        public List<IVisio.Page> GetPagesByName(string Name)
        {
            var application = this._client.Application.Get();
            var active_document = application.ActiveDocument;
            if (Name == null || Name == "*")
            {
                // return all pages
                var pages = active_document.Pages.ToEnumerable().ToList();
                return pages;
            }
            else
            {
                // return the named page
                var pages = active_document.Pages.ToEnumerable();
                var pages2= WildcardHelper.FilterObjectsByNames(pages, new[] { Name }, p => p.Name, true, WildcardHelper.FilterAction.Include).ToList();
                return pages2;
            }
        }

        public List<IVisio.Shape> GetShapes()
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var page = this._client.Page.Get();
            var shapes = page.Shapes.ToEnumerable().ToList();
            return shapes;
        }

        public List<short> GetShapeIDs()
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var page = this._client.Page.Get();
            var shapes = page.Shapes.ToEnumerable().Select(s=>s.ID16).ToList();
            return shapes;
        }
    }
}
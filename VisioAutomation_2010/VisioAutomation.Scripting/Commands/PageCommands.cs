using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class PageCommands : CommandSet
    {
        public PageCommands(Client client) :
            base(client)
        {

        }

        public void Set(IVisio.Page page)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var app = this.Client.VisioApplication;
            this.Client.WriteVerbose("Setting Active Page to \"{0}\"", page.Name);
            var window = app.ActiveWindow;
            window.Page = page;
        }

        public void Set(string name)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var doc = this.Client.VisioApplication.ActiveDocument;
            this.Client.WriteVerbose("Retrieving Page \"{0}\"", name);
            var pages = doc.Pages;
            var page = pages[name];
            this.Set(page);
        }

        public void Set(int pagenumber)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var doc = this.Client.VisioApplication.ActiveDocument;
            this.Client.WriteVerbose("Retrieving Page Number \"{0}\"", pagenumber);
            var pages = doc.Pages;
            var page = pages[pagenumber];
            this.Set(page);
        }
        
        public IVisio.Page Get()
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.VisioApplication;
            return application.ActivePage;
        }

        public void Delete(IList<IVisio.Page> pages, bool renumber)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (pages == null)
            {
                throw new System.ArgumentNullException("pages");
            }

            foreach (var page in pages)
            {
                page.Delete(renumber ? (short) 1 : (short) 0);
            }
        }

        public void Delete(IList<string> names, bool renumber)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (names == null)
            {
                throw new System.ArgumentNullException("names");
            }

            foreach (var name in names)
            {
                var app = this.Client.VisioApplication;
                var doc = app.ActiveDocument;
                var pages = doc.Pages;

                this.Client.WriteVerbose("Retrieving Page for name \"{0}\"",name);
                var page = pages.ItemU[name];
                page.Delete(renumber ? (short)1 : (short)0);
            }
        }

        public VA.Drawing.Size GetSize()
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.VisioApplication;
            var active_page = application.ActivePage;


            var query = new VA.ShapeSheet.Query.CellQuery();
            var col_height = query.AddCell(VA.ShapeSheet.SRCConstants.PageHeight);
            var col_width = query.AddCell(VA.ShapeSheet.SRCConstants.PageWidth);
            var results = query.GetResults<double>(active_page.PageSheet);
            double height = results[col_height];
            double width = results[col_width];
            var s = new VA.Drawing.Size(width, height);
            return s;
        }

        public void SetName(string name)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (name == null)
            {
                throw new System.ArgumentNullException("name");
            }

            if (name.Length < 1)
            {
                throw new System.ArgumentException("name must have at least one character");
            }

            var page = Get();
            page.NameU = name;
        }

        public IVisio.Page New(VA.Drawing.Size? size, bool isbackgroundpage)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.VisioApplication;
            var active_document = application.ActiveDocument;
            var pages = active_document.Pages;
            IVisio.Page page = pages.Add();

            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication,"New Page"))
            {
                if (size.HasValue)
                {
                    this.Client.WriteVerbose("Setting page size to {0}", size.Value);
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
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (background_page_name == null)
            {
                throw new System.ArgumentNullException("background_page_name");
            }

            var application = this.Client.VisioApplication;
            var active_document = application.ActiveDocument;
            var pages = active_document.Pages;
            var names = new HashSet<string>(pages.GetNamesU());
            if (!names.Contains(background_page_name))
            {
                string msg = string.Format("Could not find page with name \"{0}\"", background_page_name);
                throw new VA.Scripting.VisioOperationException(msg);
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
                string msg = string.Format("Cannot set page as its own background page");
                throw new VA.Scripting.VisioOperationException(msg);
            }

            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication,"Set Background Page"))
            {
                fgpage.BackPage = bgpage;
            }
        }

        public IVisio.Page Duplicate()
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication,"Duplicate Page"))
            {
                var doc = application.ActiveDocument;
                var pages = doc.Pages;
                var src_page = application.ActivePage;
                var new_page = pages.Add();

                var win = application.ActiveWindow;
                win.Page = src_page;
                VA.Pages.PageHelper.Duplicate(src_page, new_page);
                win.Page = new_page;
                return new_page;
            }
        }

        public IVisio.Page Duplicate(IVisio.Document dest_doc)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (dest_doc==null)
            {
                throw new System.ArgumentNullException("dest_doc");
            }

            var application = this.Client.VisioApplication;

            if (application.ActiveDocument == dest_doc)
            {
                throw new VA.Scripting.VisioOperationException("dest doc is same as active doc");
            }

            var src_page = application.ActivePage;
            var dest_pages = dest_doc.Pages;
            var dest_page = dest_pages[1];
            VA.Pages.PageHelper.Duplicate(src_page, dest_page);

            // the active window will be to the new document
            var active_window = application.ActiveWindow;
            //active_window.Page = dest_page;

            return dest_page;
        }

        public VA.Pages.PrintPageOrientation GetOrientation()
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.VisioApplication;
            var active_page = application.ActivePage;
            return GetOrientation(active_page);
        }

        private static VA.Pages.PrintPageOrientation GetOrientation(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            var page_sheet = page.PageSheet;
            var src = VA.ShapeSheet.SRCConstants.PrintPageOrientation;
            var orientationcell = page_sheet.CellsSRC[src.Section, src.Row, src.Cell];
            int value = orientationcell.ResultInt[IVisio.VisUnitCodes.visNumber, 0];
            return (VA.Pages.PrintPageOrientation)value;
        }

        public void SetOrientation(VA.Pages.PrintPageOrientation orientation)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.VisioApplication;

            var active_page = application.ActivePage;

            if (orientation != VA.Pages.PrintPageOrientation.Landscape && orientation != VA.Pages.PrintPageOrientation.Portrait)
            {
                throw new System.ArgumentOutOfRangeException("orientation", "must be either Portrait or Landscape");
            }

            var old_orientation = GetOrientation(active_page);

            if (old_orientation == orientation)
            {
                // don't need to do anything
                return;
            }

            var old_size = this.GetSize();

            double new_height = old_size.Width;
            double new_width = old_size.Height;

            var update = new VA.ShapeSheet.Update(3);
            update.SetFormula(VA.ShapeSheet.SRCConstants.PageWidth, new_width);
            update.SetFormula(VA.ShapeSheet.SRCConstants.PageHeight, new_height);
            update.SetFormula(VA.ShapeSheet.SRCConstants.PrintPageOrientation, (int)orientation);

            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication,"Set Page Orientation"))
            {
                update.Execute(active_page.PageSheet);
            }
        }



        public void ResizeToFitContents(VA.Drawing.Size bordersize, bool zoom_to_page)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication,"Resize Page to Fit Contents"))
            {
                var active_page = application.ActivePage;
                active_page.ResizeToFitContents(bordersize);
                if (zoom_to_page)
                {
                    this.Client.View.Zoom(VA.Scripting.Zoom.ToPage);
                }
            }
        }

        public void ResetOrigin(IVisio.Page page)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.VisioApplication;
            if (page == null)
            {
                page = application.ActivePage;
            }

            var update = new VA.ShapeSheet.Update();

            update.SetFormula(VA.ShapeSheet.SRCConstants.XGridOrigin, "0.0");
            update.SetFormula(VA.ShapeSheet.SRCConstants.YGridOrigin, "0.0");
            update.SetFormula(VA.ShapeSheet.SRCConstants.XRulerOrigin, "0.0");
            update.SetFormula(VA.ShapeSheet.SRCConstants.YRulerOrigin, "0.0");

            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication,"Reset Page Origin"))
            {
                update.Execute(page.PageSheet);
            }
        }

        public void SetSize(VA.Drawing.Size new_size)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication,"Set Page Size"))
            {
                var active_page = application.ActivePage;
                var page_sheet = active_page.PageSheet;
                var update = new VA.ShapeSheet.Update(2);
                update.SetFormula(VA.ShapeSheet.SRCConstants.PageWidth, new_size.Width);
                update.SetFormula(VA.ShapeSheet.SRCConstants.PageHeight, new_size.Height);
                update.Execute(page_sheet);
            }
        }

        public void SetSize(double? width, double? height)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (!width.HasValue && !height.HasValue)
            {
                // nothing to do
                return;
            }

            var old_size = this.GetSize();
            var w = width.GetValueOrDefault(old_size.Width);
            var h = height.GetValueOrDefault(old_size.Height);
            var new_size = new VA.Drawing.Size(w, h);
            SetSize(new_size);
        }

        public void SetHeight(double height)
        {
            SetSize(null, height);
        }

        public void SetWidth(double width)
        {
            SetSize(width, null);
        }

        public void GoTo(VA.Scripting.PageDirection flags)
        {
            this.Client.Application.AssertApplicationAvailable();

            var application = this.Client.VisioApplication;
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
            this.Client.Application.AssertApplicationAvailable();

            if (pages == null)
            {
                throw new System.ArgumentNullException("pages");
            }

            var app = pages.Application;
            var active_document = app.ActiveDocument;
            if (pages.Document != active_document)
            {
                throw new VA.Scripting.VisioOperationException("Page.Document is not application's ActiveDocument");
            }

            if (pages.Count < 2)
            {
                throw new VA.Scripting.VisioOperationException("Only 1 page available. Navigation not possible.");
            }

            var activepage = app.ActivePage;

            int cur_index = activepage.Index;
            const int min_index = 1;
            int max_index = pages.Count;
            int new_index = move_in_range(cur_index, min_index, max_index, flags);
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
                throw new System.ArgumentOutOfRangeException("max");
            }

            if (cur < min)
            {
                throw new System.ArgumentOutOfRangeException("cur");
            }

            if (cur > max)
            {
                throw new System.ArgumentOutOfRangeException("cur");
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
                throw new System.ArgumentOutOfRangeException("direction");
            }
        }

        public IList<IVisio.Shape> GetShapesByID(int[] shapeids)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var page = this.Client.Page.Get();
            var shapes = page.Shapes;
            var shapes_list = new List<IVisio.Shape>(shapeids.Length);
            foreach (int id in shapeids)
            {
                var shape = shapes.ItemFromID[id];
                shapes_list.Add(shape);
            }
            return shapes_list;
        }

        public IList<IVisio.Shape> GetShapesByName(string[] shapenames)
        {
            return this.GetShapesByName(shapenames, false);
        }

        public IList<IVisio.Shape> GetShapesByName(string[] shapenames, bool ignore_bad_names)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var page = this.Client.Page.Get();
            var shapes = page.Shapes;

            var cached_shapes_list = new List<IVisio.Shape>(shapes.Count);
            cached_shapes_list.AddRange(shapes.AsEnumerable());
            
            if (shapenames.Contains("*"))
            {
                // if any of the shape names contains a simple wildcard then return all the shapes
                return cached_shapes_list;
            }

            // otherwise we start checking for each name
            var shapes_list = VA.TextUtil.FilterObjectsByNames(cached_shapes_list, shapenames, s => s.Name, true, VA.TextUtil.FilterAction.Include).ToList();

            return shapes_list;
        }

        public IList<IVisio.Page> GetPagesByName(string Name)
        {
            var active_document = this.Client.VisioApplication.ActiveDocument;
            if (Name == null || Name == "*")
            {
                // return all pages
                var pages = active_document.Pages.AsEnumerable().ToList();
                return pages;
            }
            else
            {
                // return the named page
                var pages = active_document.Pages.AsEnumerable();
                var pages2= VA.TextUtil.FilterObjectsByNames(pages, new[] { Name }, p => p.Name, true, VA.TextUtil.FilterAction.Include).ToList();
                return pages2;
            }
        }
    }
}
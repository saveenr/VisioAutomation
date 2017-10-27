using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class PageCommands : CommandSet
    {
        internal PageCommands(Client client) :
            base(client)
        {

        }

        public void SetActivePage(IVisio.Page page)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();
            var app = cmdtarget.Application;
            this._client.Output.WriteVerbose("Setting Active Page to \"{0}\"", page.Name);
            var window = app.ActiveWindow;
            window.Page = page;
        }

        public void SetActivePageByPageName(string name)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();
            var doc = cmdtarget.ActiveDocument;
            this._client.Output.WriteVerbose("Retrieving Page \"{0}\"", name);
            var pages = doc.Pages;
            var page = pages[name];
            this.SetActivePage(page);
        }

        public void SetActivePageByPageNumber(int pagenumber)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();


            var application = cmdtarget.Application;
            var doc = application.ActiveDocument;
            this._client.Output.WriteVerbose("Retrieving Page Number \"{0}\"", pagenumber);
            var pages = doc.Pages;
            var page = pages[pagenumber];
            this.SetActivePage(page);
        }
        
        public IVisio.Page GetActivePage()
        {
            var cmdtarget = this._client.GetCommandTargetDocument();

            var application = cmdtarget.Application;
            return application.ActivePage;
        }

        public void DeletePages(IList<IVisio.Page> pages, bool renumber)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();

            if (pages == null)
            {
                throw new System.ArgumentNullException(nameof(pages));
            }

            foreach (var page in pages)
            {
                page.Delete(renumber ? (short) 1 : (short) 0);
            }
        }

        public VisioAutomation.Geometry.Size GetActivePageSize()
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            var active_page = cmdtarget.ActivePage;

            var query = new VisioAutomation.ShapeSheet.Query.CellQuery();
            var col_height = query.Columns.Add(VisioAutomation.ShapeSheet.SrcConstants.PageHeight, nameof(VisioAutomation.ShapeSheet.SrcConstants.PageHeight));
            var col_width = query.Columns.Add(VisioAutomation.ShapeSheet.SrcConstants.PageWidth, nameof(VisioAutomation.ShapeSheet.SrcConstants.PageWidth));

            var results = query.GetResults<double>(active_page.PageSheet);
            double height = results.Cells[col_height];
            double width = results.Cells[col_width];
            var s = new VisioAutomation.Geometry.Size(width, height);
            return s;
        }

        public void SetActivePageName(string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            if (name.Length < 1)
            {
                throw new System.ArgumentException("name cannot be empty", nameof(name));
            }

            var cmdtarget = this._client.GetCommandTargetPage();
            cmdtarget.ActivePage.NameU = name;
        }

        public IVisio.Page NewPage(VisioAutomation.Geometry.Size? size, bool isbackgroundpage)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();

            var active_document = cmdtarget.ActiveDocument;
            var pages = active_document.Pages;
            IVisio.Page new_page;

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(NewPage)))
            {
                new_page = pages.Add();

                if (size.HasValue)
                {
                    var target_pages = new Models.TargetPages(new_page);
                    this.SetPageSize(target_pages, size.Value);
                }

                if (isbackgroundpage)
                {
                    new_page.Background = 1;
                }
            }

            return new_page;
        }

        public void SetActivePageBackground(string background_page_name)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();

            if (background_page_name == null)
            {
                throw new System.ArgumentNullException(nameof(background_page_name));
            }

            var app = cmdtarget.Application;
            var application = app;
            var active_document = application.ActiveDocument;
            var pages = active_document.Pages;
            var names = new HashSet<string>(pages.GetNamesU());
            if (!names.Contains(background_page_name))
            {
                string msg = string.Format("Could not find page with name \"{0}\"", background_page_name);
                throw new VisioAutomation.Exceptions.VisioOperationException(msg);
            }

            var bgpage = pages.ItemU[background_page_name];
            var fgpage = application.ActivePage;

            // Set the background page
            // Check that the intended background is indeed a background page
            if (bgpage.Background == 0)
            {
                string msg = string.Format("Page \"{0}\" is not a background page", bgpage.Name);
                throw new VisioAutomation.Exceptions.VisioOperationException(msg);
            }

            // don't allow the page to be set as a background to itself
            if (fgpage == bgpage)
            {
                string msg = "Cannot set page as its own background page";
                throw new VisioAutomation.Exceptions.VisioOperationException(msg);
            }

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(SetActivePageBackground)))
            {
                fgpage.BackPage = bgpage;
            }
        }

        public IVisio.Page DuplicateActivePage()
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DuplicateActivePage)))
            {
                var pages = cmdtarget.ActiveDocument.Pages;

                var src_page = cmdtarget.Application.ActivePage;
                var new_page = pages.Add();

                var win = cmdtarget.Application.ActiveWindow;
                win.Page = src_page;
                VisioAutomation.Pages.PageHelper.Duplicate(src_page, new_page);
                win.Page = new_page;
                return new_page;
            }
        }

        public IVisio.Page DuplicateActivePageToDocument(IVisio.Document dest_doc)
        {
            if (dest_doc == null)
            {
                throw new System.ArgumentNullException(nameof(dest_doc));
            }

            var cmdtarget = this._client.GetCommandTargetPage();

            if (cmdtarget.ActiveDocument == dest_doc)
            {
                throw new VisioAutomation.Exceptions.VisioOperationException("dest doc is same as active doc");
            }

            var src_page = cmdtarget.ActivePage;
            var dest_pages = dest_doc.Pages;
            var dest_page = dest_pages[1];
            VisioAutomation.Pages.PageHelper.Duplicate(src_page, dest_page);

            // the active window will be to the new document
            var active_window = cmdtarget.Application.ActiveWindow;
            //active_window.Page = dest_page;

            return dest_page;
        }

        public Models.PageOrientation GetActivePageOrientation()
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            var active_page = cmdtarget.ActivePage;
            return PageCommands._GetPageOrientation(active_page);
        }

        private static Models.PageOrientation _GetPageOrientation(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            var page_sheet = page.PageSheet;
            var src = VisioAutomation.ShapeSheet.SrcConstants.PrintPageOrientation;
            var orientationcell = page_sheet.CellsSRC[src.Section, src.Row, src.Cell];
            int value = orientationcell.ResultInt[IVisio.VisUnitCodes.visNumber, 0];
            return (Models.PageOrientation)value;
        }

        public void SetActivePageOrientation(Models.PageOrientation orientation)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            if (orientation != VisioScripting.Models.PageOrientation.Landscape && orientation != VisioScripting.Models.PageOrientation.Portrait)
            {
                throw new System.ArgumentOutOfRangeException(nameof(orientation), "must be either Portrait or Landscape");
            }

            var old_orientation = PageCommands._GetPageOrientation(cmdtarget.Application.ActivePage);

            if (old_orientation == orientation)
            {
                // don't need to do anything
                return;
            }

            var old_size = this.GetActivePageSize();

            double new_height = old_size.Width;
            double new_width = old_size.Height;

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetFormula(VisioAutomation.ShapeSheet.SrcConstants.PageWidth, new_width);
            writer.SetFormula(VisioAutomation.ShapeSheet.SrcConstants.PageHeight, new_height);
            writer.SetFormula(VisioAutomation.ShapeSheet.SrcConstants.PrintPageOrientation, (int)orientation);

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(SetActivePageOrientation)))
            {
                writer.Commit(cmdtarget.ActivePage.PageSheet);
            }
        }

        public void ResizePageToFitContents(Models.TargetPages target_pages, VisioAutomation.Geometry.Size bordersize)
        {
            var pages = target_pages.Resolve(this._client);

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(ResizePageToFitContents)))
            {
                foreach (var page in pages)
                {
                    page.ResizeToFitContents(bordersize);
                }
            }
        }

        public void SetPageFormatCells(Models.TargetPages targetpages, VisioAutomation.Pages.PageFormatCells cells)
        {
            var pages = targetpages.Resolve(this._client);

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(SetPageFormatCells)))
            {
                foreach (var page in pages)
                {
                    var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
                    cells.SetFormulas(writer);
                    writer.BlastGuards = true;
                    writer.Commit(page);
                }
            }
        }

        public void SetPageSize(Models.TargetPages targetpages, VisioAutomation.Geometry.Size new_size)
        {
            var pages = targetpages.Resolve(this._client);

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(SetPageSize)))
            {
                foreach (var page in pages)
                {
                    var page_sheet = page.PageSheet;
                    var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
                    writer.SetFormula(VisioAutomation.ShapeSheet.SrcConstants.PageWidth, new_size.Width);
                    writer.SetFormula(VisioAutomation.ShapeSheet.SrcConstants.PageHeight, new_size.Height);
                    writer.Commit(page_sheet);
                }
            }
        }

        public void SetActivePageSize(double? width, double? height)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            if (!width.HasValue && !height.HasValue)
            {
                // nothing to do
                return;
            }

            var old_size = this.GetActivePageSize();
            var w = width.GetValueOrDefault(old_size.Width);
            var h = height.GetValueOrDefault(old_size.Height);
            var new_size = new VisioAutomation.Geometry.Size(w, h);
            this.SetPageSize(new Models.TargetPages(cmdtarget.ActivePage),new_size);
        }

        public void SetActivePageByDirection(Models.PageDirection flags)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            var docpages = cmdtarget.ActiveDocument.Pages;
            if (docpages.Count < 2)
            {
                return;
            }

            var pages = docpages;
            this._GoTo(pages, flags, cmdtarget);
        }

        private void _GoTo(IVisio.Pages pages, Models.PageDirection flags, CommandTarget cmdtarget)
        {
            if (pages == null)
            {
                throw new System.ArgumentNullException(nameof(pages));
            }

            if (pages.Count < 2)
            {
                throw new VisioAutomation.Exceptions.VisioOperationException("Only 1 page available. Navigation not possible.");
            }

            int cur_index = cmdtarget.ActivePage.Index;
            const int min_index = 1;
            int max_index = pages.Count;
            int new_index = PageCommands.move_in_range(cur_index, min_index, max_index, flags);
            if (cur_index != new_index)
            {
                var doc_pages = cmdtarget.ActiveDocument.Pages;
                var page = doc_pages[new_index];

                var active_window = cmdtarget.Application.ActiveWindow;
                active_window.Page = page;
            }
        }

        internal static int move_in_range(int cur, int min, int max, Models.PageDirection direction)
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

            switch (direction)
            {
                case VisioScripting.Models.PageDirection.Next:
                    return System.Math.Min(cur + 1, max);
                case VisioScripting.Models.PageDirection.Previous:
                    return System.Math.Max(cur - 1, min);
                case VisioScripting.Models.PageDirection.First:
                    return min;
                case VisioScripting.Models.PageDirection.Last:
                    return max;
                default:
                    throw new System.ArgumentOutOfRangeException(nameof(direction));
            }
        }

        public List<IVisio.Shape> GetShapesOnActivePageByID(int[] shapeids)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            var shapes = cmdtarget.ActivePage.Shapes;
            var shapes_list = new List<IVisio.Shape>(shapeids.Length);
            foreach (int id in shapeids)
            {
                var shape = shapes.ItemFromID[id];
                shapes_list.Add(shape);
            }
            return shapes_list;
        }

        public List<IVisio.Shape> GetShapesOnActivePageByName(string[] shapenames)
        {
            return this.GetShapesOnActivePageByName(shapenames, false);
        }

        public List<IVisio.Shape> GetShapesOnActivePageByName(string[] shapenames, bool ignore_bad_names)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();
            var shapes = cmdtarget.ActivePage.Shapes;
            var cached_shapes_list = new List<IVisio.Shape>(shapes.Count);
            cached_shapes_list.AddRange(shapes.ToEnumerable());
            
            if (shapenames.Contains("*"))
            {
                // if any of the shape names contains a simple wildcard then return all the shapes
                return cached_shapes_list;
            }

            // otherwise we start checking for each name
            var shapes_list = VisioScripting.Helpers.WildcardHelper.FilterObjectsByNames(cached_shapes_list, shapenames, s => s.Name, true, VisioScripting.Helpers.WildcardHelper.FilterAction.Include).ToList();

            return shapes_list;
        }

        public List<IVisio.Page> FindPagesInActiveDocumentByName(string name, Models.PageType pagetype)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();

            var active_document = cmdtarget.ActiveDocument;
            if (VisioScripting.Helpers.WildcardHelper.NullOrStar(name))
            {
                // return all pages
                var all_pages = active_document.Pages.ToList();
                all_pages = filter_pages_by_type(all_pages, pagetype);
                return all_pages;
            }
            else
            {
                // return the named page
                var all_pages = active_document.Pages.ToEnumerable();
                var named_pages= VisioScripting.Helpers.WildcardHelper.FilterObjectsByNames(all_pages, new[] { name }, p => p.Name, true, VisioScripting.Helpers.WildcardHelper.FilterAction.Include).ToList();
                named_pages = filter_pages_by_type(named_pages, pagetype);

                return named_pages;
            }
        }

        private List<IVisio.Page> filter_pages_by_type(List<IVisio.Page> pages, Models.PageType pagetype)
        {
            if (pages == null)
            {
                return null;
            }

            if (pagetype == Models.PageType.Any)
            {
                return pages;
            }
            if (pagetype == Models.PageType.Foreground)
            {
                return pages.Where(p=>p.Background==0).ToList();
            }

            if (pagetype == Models.PageType.Background)
            {
                return pages.Where(p => p.Background != 0).ToList();
            }

            string msg = "Unsupported value for pagetype";
            throw new System.ArgumentOutOfRangeException(nameof(pagetype),msg);
        }

        public List<IVisio.Shape> GetShapesOnActivePage()
        {
            var cmdtarget = this._client.GetCommandTargetPage();
            var shapes = cmdtarget.ActivePage.Shapes.ToList();
            return shapes;
        }
    }
}
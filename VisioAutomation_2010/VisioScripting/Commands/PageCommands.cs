using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet;
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
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.RequireDocument);
            var app = cmdtarget.Application;
            this._client.Output.WriteVerbose("Setting Active Page to \"{0}\"", page.Name);
            var window = app.ActiveWindow;
            window.Page = page;
        }

        
        public IVisio.Page GetActivePage()
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.RequireDocument);
            var application = cmdtarget.Application;
            return application.ActivePage;
        }

        public void DeletePages(TargetPages targetpages, bool renumber)
        {
            targetpages = targetpages.ResolveToPages(this._client);

            foreach (var page in targetpages.Pages)
            {
                page.Delete(renumber ? (short) 1 : (short) 0);
            }
        }

        public List<VisioAutomation.Geometry.Size> GetPageSize(TargetPages targetpages)
        {
            targetpages = targetpages.ResolveToPages(this._client);

            if (targetpages.Pages.Count < 1)
            {
                return  new List<VisioAutomation.Geometry.Size>(0);
            }

            var sizes = new List<VisioAutomation.Geometry.Size>(targetpages.Pages.Count);

            foreach (var page in targetpages.Pages)
            {
                var size = VisioAutomation.Pages.PageHelper.GetSize(page);
                sizes.Add(size);
            }

            return sizes;
        }

        public IVisio.Page NewPage(VisioScripting.TargetDocument targetdoc, VisioAutomation.Geometry.Size? size, bool isbackgroundpage)
        {
            targetdoc = targetdoc.ResolveToDocument(this._client);
            var pages = targetdoc.Document.Pages;
            IVisio.Page new_page;

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(NewPage)))
            {
                new_page = pages.Add();

                if (size.HasValue)
                {
                    var targetpages = new TargetPages(new_page);
                    this.SetPageSize(targetpages, size.Value);
                }

                if (isbackgroundpage)
                {
                    new_page.Background = 1;
                }
            }

            return new_page;
        }

        public void SetPageBackground(TargetPages targetpages, string background_page_name)
        {
            if (background_page_name == null)
            {
                throw new System.ArgumentNullException(nameof(background_page_name));
            }

            targetpages = targetpages.ResolveToPages(this._client);

            if (targetpages.Pages.Count < 1)
            {
                return;
            }

            var page0 = targetpages.Pages[0];
            var doc = page0.Document;
            var doc_pages = doc.Pages;

            var names = new HashSet<string>(doc_pages.GetNamesU());
            if (!names.Contains(background_page_name))
            {
                string msg = string.Format("Could not find page with name \"{0}\"", background_page_name);
                throw new VisioAutomation.Exceptions.VisioOperationException(msg);
            }

            var bgpage = doc_pages.ItemU[background_page_name];

            // Set the background page
            // Check that the intended background is indeed a background page
            if (bgpage.Background == 0)
            {
                string msg = string.Format("Page \"{0}\" is not a background page", bgpage.Name);
                throw new VisioAutomation.Exceptions.VisioOperationException(msg);
            }

            // don't allow the page to be set as a background to itself

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(SetPageBackground)))
            {
                foreach (var page in targetpages.Pages)
                {
                    if (page == bgpage)
                    {
                        string msg = "Cannot set page as its own background page";
                        throw new VisioAutomation.Exceptions.VisioOperationException(msg);
                    }

                    page.BackPage = bgpage;
                }
            }
        }

        public IVisio.Page DuplicatePage(TargetPage targetpage)
        {
            targetpage = targetpage.ResolveToPage(this._client);

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DuplicatePage)))
            {
                var src_page = targetpage.Page;
                var doc = src_page.Document;
                var pages = doc.Pages;
                var new_page = pages.Add();
                var app = doc.Application;

                var win = app.ActiveWindow;
                win.Page = src_page;
                VisioAutomation.Pages.PageHelper.Duplicate(src_page, new_page);
                win.Page = new_page;
                return new_page;
            }
        }

        public IVisio.Page DuplicatePageToDocument(TargetPage targetpage, IVisio.Document dest_doc)
        {
            targetpage = targetpage.ResolveToPage(this._client);

            if (dest_doc == null)
            {
                throw new System.ArgumentNullException(nameof(dest_doc));
            }

            if (targetpage.Page.Document == dest_doc)
            {
                throw new VisioAutomation.Exceptions.VisioOperationException("dest doc is same as pages src doc");
            }

            var dest_pages = dest_doc.Pages;
            var dest_page = dest_pages.Add();
            VisioAutomation.Pages.PageHelper.Duplicate(targetpage.Page, dest_page);

            return dest_page;
        }

        public Models.PageOrientation GetPageOrientation( TargetPage targetpage )
        {
            targetpage = targetpage.ResolveToPage(this._client);
            return PageCommands._get_page_orientation(targetpage.Page);
        }
        
        private static Models.PageOrientation _get_page_orientation(IVisio.Page page)
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

        public void SetPageOrientation(TargetPages targetpages, Models.PageOrientation orientation)
        {
            if (orientation != VisioScripting.Models.PageOrientation.Landscape && orientation != VisioScripting.Models.PageOrientation.Portrait)
            {
                throw new System.ArgumentOutOfRangeException(nameof(orientation), "must be either Portrait or Landscape");
            }

            targetpages = targetpages.ResolveToPages(this._client);

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(SetPageOrientation)))
            {

                foreach (var page in targetpages.Pages)
                {
                    var old_orientation = PageCommands._get_page_orientation(page);

                    if (old_orientation == orientation)
                    {
                        // don't need to do anything
                        return;
                    }

                    var old_size = VisioAutomation.Pages.PageHelper.GetSize(page);

                    double new_height = old_size.Width;
                    double new_width = old_size.Height;

                    var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
                    writer.SetValue(VisioAutomation.ShapeSheet.SrcConstants.PageWidth, new_width);
                    writer.SetValue(VisioAutomation.ShapeSheet.SrcConstants.PageHeight, new_height);
                    writer.SetValue(VisioAutomation.ShapeSheet.SrcConstants.PrintPageOrientation, (int)orientation);

                    writer.Commit(page.PageSheet, CellValueType.Formula);
                }

            }

        }
        public void ResizePageToFitContents(TargetPages targetpages, VisioAutomation.Geometry.Size bordersize)
        {
            targetpages = targetpages.ResolveToPages(this._client);

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(ResizePageToFitContents)))
            {
                foreach (var page in targetpages.Pages)
                {
                    page.ResizeToFitContents(bordersize);
                }
            }
        }

        public void SetPageFormatCells(TargetPages targetpages, VisioAutomation.Pages.PageFormatCells cells)
        {
            targetpages = targetpages.ResolveToPages(this._client);

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(SetPageFormatCells)))
            {
                foreach (var page in targetpages.Pages)
                {
                    var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
                    writer.SetValues(cells);
                    writer.BlastGuards = true;
                    writer.Commit(page, CellValueType.Formula);
                }
            }
        }

        public void SetPageSize(TargetPages targetpages, VisioAutomation.Geometry.Size new_size)
        {
            targetpages = targetpages.ResolveToPages(this._client);

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(SetPageSize)))
            {
                foreach (var page in targetpages.Pages)
                {
                    var page_sheet = page.PageSheet;
                    var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
                    writer.SetValue(VisioAutomation.ShapeSheet.SrcConstants.PageWidth, new_size.Width);
                    writer.SetValue(VisioAutomation.ShapeSheet.SrcConstants.PageHeight, new_size.Height);
                    writer.Commit(page_sheet, CellValueType.Formula);
                }
            }
        }

        public void SetPageSize(TargetPage targetpage, double? width, double? height)
        {
            if (!width.HasValue && !height.HasValue)
            {
                // nothing to do
                return;
            }

            targetpage = targetpage.ResolveToPage(this._client);

            var old_size = VisioAutomation.Pages.PageHelper.GetSize(targetpage.Page);
            var w = width.GetValueOrDefault(old_size.Width);
            var h = height.GetValueOrDefault(old_size.Height);
            var new_size = new VisioAutomation.Geometry.Size(w, h);
            this.SetPageSize(new TargetPages(targetpage.Page),new_size);
        }

        public void SetActivePage(VisioScripting.TargetDocument targetdoc, Models.PageRelativePosition flags)
        {
            targetdoc = targetdoc.ResolveToDocument(this._client);

            var docpages = targetdoc.Document.Pages;
            if (docpages.Count < 2)
            {
                return;
            }

            var pages = docpages;

            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.RequirePage);

            this._go_to_page(pages, flags, cmdtarget);
        }

        public void LayoutPage(TargetPages targetpages, VisioAutomation.Models.LayoutStyles.LayoutStyleBase layout)
        {
            targetpages = targetpages.ResolveToPages(this._client);

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(SetPageSize)))
            {
                foreach (var page in targetpages.Pages)
                {
                    layout.Apply(page);
                }
            }
        }

        private void _go_to_page(IVisio.Pages pages, Models.PageRelativePosition flags, VisioScripting.CommandTarget cmdtarget)
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

        internal static int move_in_range(int cur, int min, int max, Models.PageRelativePosition relative_position)
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

            return relative_position switch
            {
                VisioScripting.Models.PageRelativePosition.Next => System.Math.Min(cur + 1, max),
                VisioScripting.Models.PageRelativePosition.Previous => System.Math.Max(cur - 1, min),
                VisioScripting.Models.PageRelativePosition.First => min,
                VisioScripting.Models.PageRelativePosition.Last => max,
                _ => throw new System.ArgumentOutOfRangeException(nameof(relative_position)),
            };
        }

        public List<IVisio.Shape> GetShapesOnPageByID(TargetPage targetpage, int[] shapeids)
        {
            targetpage = targetpage.ResolveToPage(this._client);
            var shapes = targetpage.Page.Shapes;
            var shapes_list = new List<IVisio.Shape>(shapeids.Length);
            foreach (int id in shapeids)
            {
                var shape = shapes.ItemFromID[id];
                shapes_list.Add(shape);
            }
            return shapes_list;
        }

        public List<IVisio.Shape> GetShapesOnPageByName(TargetPage targetpage, string[] names)
        {
            targetpage = targetpage.ResolveToPage(this._client);

            var cached_shapes_list = targetpage.Page.Shapes.ToList();
            
            if (names.Contains("*"))
            {
                // if any of the shape names contains a simple wildcard then return all the shapes
                return cached_shapes_list;
            }

            // otherwise we start checking for each name
            var shapes_list = VisioScripting.Helpers.WildcardHelper.FilterObjectsByNames(cached_shapes_list, names, s => s.Name, true, VisioScripting.Helpers.WildcardHelper.FilterAction.Include).ToList();

            return shapes_list;
        }

        public List<IVisio.Page> FindPagesInDocument(TargetDocument targetdoc, string name)
        {
            targetdoc = targetdoc.ResolveToDocument(this._client);

            if (VisioScripting.Helpers.WildcardHelper.NullOrStar(name))
            {
                // return all pages
                var all_pages = targetdoc.Document.Pages.ToList();
                return all_pages;
            }
            else
            {
                // return the named page
                var all_pages = targetdoc.Document.Pages.ToEnumerable();
                var named_pages = VisioScripting.Helpers.WildcardHelper.FilterObjectsByNames(all_pages, new[] {name},
                    p => p.Name, true, VisioScripting.Helpers.WildcardHelper.FilterAction.Include).ToList();
                return named_pages;
            }
        }
        
        public List<IVisio.Shape> GetShapesOnPage(TargetPage targetpage)
        {
            targetpage = targetpage.ResolveToPage(this._client);
            var shapes = targetpage.Page.Shapes.ToList();
            return shapes;
        }
    }
}
using System;
using System.Collections.Generic;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class PageCommands : SessionCommands
    {
        public PageCommands(Session session) :
            base(session)
        {

        }

        public IVisio.Page GetPage()
        {
            if (!this.Session.HasActiveDrawing)
            {
                throw new AutomationException("No Drawing available");
            }

            var application = this.Session.VisioApplication;
            return application.ActivePage;
        }

        public VA.Drawing.Size GetPageSize()
        {
            if (!this.Session.HasActiveDrawing)
            {
                throw new AutomationException("No Drawing available");
            }

            var application = this.Session.VisioApplication;
            var active_page = application.ActivePage;
            return active_page.GetSize();
        }

        [Obsolete]
        public string GetPageName()
        {
            return GetPage().NameU;
        }

        public void SetPageName(string name)
        {
            if (name == null)
            {
                throw new ArgumentNullException("name");
            }

            if (name.Length < 1)
            {
                throw new ArgumentException("name must have at least one character");
            }

            var application = this.Session.VisioApplication;
            var active_document = application.ActiveDocument;
            var pages = active_document.Pages;
            var pagenames = new HashSet<string>(pages.GetNamesU());
            if (pagenames.Contains(name))
            {
                throw new AutomationException("Page already exists with this name");
            }

            var page = GetPage();
            page.NameU = name;
        }

        public IVisio.Page NewPage(VA.Drawing.Size? size, bool isbackgroundpage)
        {
            IVisio.Page page;
            var application = this.Session.VisioApplication;
            var active_document = application.ActiveDocument;
            var pages = active_document.Pages;
            page = pages.Add();

            using (var undoscope = application.CreateUndoScope())
            {
                if (size.HasValue)
                {
                    this.Session.Write(OutputStream.Verbose,"Setting page size to {0}", size.Value);
                    page.SetSize(size.Value);
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
            if (background_page_name == null)
            {
                throw new ArgumentNullException("background_page_name");
            }

            if (!this.Session.HasActiveDrawing)
            {
                return;
            }

            var application = this.Session.VisioApplication;
            var active_document = application.ActiveDocument;
            var pages = active_document.Pages;
            var names = new HashSet<string>(pages.GetNamesU());
            if (!names.Contains(background_page_name))
            {
                string msg = String.Format("Could not find page with name \"{0}\"", background_page_name);
                throw new AutomationException(msg);
            }

            var bgpage = pages.ItemU[background_page_name];
            var fgpage = application.ActivePage;

            using (var undoscope = application.CreateUndoScope())
            {
                VA.PageHelper.SetBackgroundPage(fgpage, bgpage);
            }
        }

        public void DuplicatePage()
        {
            if (!this.Session.HasActiveDrawing)
            {
                return;
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                VA.PageHelper.Duplicate(application.ActivePage);
            }
        }

        public void DuplicatePageToNewDocument()
        {
            if (!this.Session.HasActiveDrawing)
            {
                return;
            }

            var application = this.Session.VisioApplication;
            var active_page = application.ActivePage;
            var page_to_dupe = active_page;
            var documents = application.Documents;
            var dest_doc = documents.Add(String.Empty);
            page_to_dupe.Activate();
            string page_name = page_to_dupe.NameU;
            var destpages = dest_doc.Pages;
            var dest_page = destpages[1];
            VA.PageHelper.DuplicateToDocument(active_page, dest_doc, dest_page, page_name, true);
            dest_doc.Activate();
            dest_page.Activate();
        }

        public VA.Layout.PrintPageOrientation GetPageOrientation()
        {
            if (!this.Session.HasActiveDrawing)
            {
                throw new AutomationException("No active page");
            }

            var application = this.Session.VisioApplication;
            var active_page = application.ActivePage;
            return VA.PageHelper.GetOrientation(active_page);
        }

        public void SetPageOrientation(VA.Layout.PrintPageOrientation orientation)
        {
            if (!this.Session.HasActiveDrawing)
            {
                return;
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                var active_page = application.ActivePage;
                VA.PageHelper.SetOrientation(active_page, orientation);
            }
        }

        public void ResizeToFitContents(VA.Drawing.Size bordersize, bool zoom_to_page)
        {
            if (!this.Session.HasActiveDrawing)
            {
                return;
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                var active_page = application.ActivePage;
                active_page.ResizeToFitContents(bordersize);
                if (zoom_to_page)
                {
                    this.Session.View.Zoom(VA.Scripting.Zoom.ToPage);
                }
            }
        }

        public void ResetPageOrigin()
        {
            if (!this.Session.HasActiveDrawing)
            {
                return;
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                var active_page = application.ActivePage;
                VA.PageHelper.ResetOrigin(active_page);
            }
        }

        public void SetPageSize(VA.Drawing.Size new_size)
        {
            if (!this.Session.HasActiveDrawing)
            {
                return;
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                var active_page = application.ActivePage;
                active_page.SetSize(new_size);
            }
        }

        public void SetPageSize(double? width, double? height)
        {
            if (!this.Session.HasActiveDrawing)
            {
                return;
            }

            if (!width.HasValue && !height.HasValue)
            {
                // nothing to do
                return;
            }

            var application = this.Session.VisioApplication;
            var active_page = application.ActivePage;
            var old_size = active_page.GetSize();
            var w = width.GetValueOrDefault(old_size.Width);
            var h = height.GetValueOrDefault(old_size.Height);
            var new_size = new VA.Drawing.Size(w, h);
            SetPageSize(new_size);
        }

        public void SetPageHeight(double height)
        {
            SetPageSize(null, height);
        }

        public void SetPageWidth(double width)
        {
            SetPageSize(width, null);
        }

        public void NavigateToPage(PageNavigation flags)
        {
            var application = this.Session.VisioApplication;
            var active_document = application.ActiveDocument;
            var docpages = active_document.Pages;
            if (docpages.Count < 2)
            {
                return;
            }

            var pages = docpages;
            VA.PageHelper.NavigateTo(pages, flags);
        }
    }
}
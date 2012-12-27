using System.Collections.Generic;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class PageCommands : CommandSet
    {
        public PageCommands(Session session) :
            base(session)
        {

        }

        public IVisio.Page Get()
        {
            if (!this.Session.HasActiveDrawing)
            {
                throw new AutomationException("No Drawing available");
            }

            var application = this.Session.VisioApplication;
            return application.ActivePage;
        }

        public VA.Drawing.Size GetSize()
        {
            if (!this.Session.HasActiveDrawing)
            {
                throw new AutomationException("No Drawing available");
            }

            var application = this.Session.VisioApplication;
            var active_page = application.ActivePage;
            return VA.Pages.PageHelper.GetSize(active_page);
        }

        public void SetName(string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException("name");
            }

            if (name.Length < 1)
            {
                throw new System.ArgumentException("name must have at least one character");
            }

            var application = this.Session.VisioApplication;
            var active_document = application.ActiveDocument;
            var pages = active_document.Pages;
            var pagenames = new HashSet<string>(pages.GetNamesU());
            if (pagenames.Contains(name))
            {
                throw new AutomationException("Page already exists with this name");
            }

            var page = Get();
            page.NameU = name;
        }

        public IVisio.Page New(VA.Drawing.Size? size, bool isbackgroundpage)
        {
            IVisio.Page page;
            var application = this.Session.VisioApplication;
            var active_document = application.ActiveDocument;
            var pages = active_document.Pages;
            page = pages.Add();

            using (var undoscope = application.CreateUndoScope("New Page"))
            {
                if (size.HasValue)
                {
                    this.Session.Write(OutputStream.Verbose,"Setting page size to {0}", size.Value);
                    VA.Pages.PageHelper.SetSize(page,size.Value);
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
                throw new System.ArgumentNullException("background_page_name");
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
                string msg = string.Format("Could not find page with name \"{0}\"", background_page_name);
                throw new AutomationException(msg);
            }

            var bgpage = pages.ItemU[background_page_name];
            var fgpage = application.ActivePage;

            using (var undoscope = application.CreateUndoScope("Set Background Page"))
            {
                VA.Pages.PageHelper.SetBackgroundPage(fgpage, bgpage);
            }
        }

        public void Duplicate(string dest_pagename)
        {
            if (!this.Session.HasActiveDrawing)
            {
                return;
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope("Duplicate Page"))
            {
                VA.Pages.PageHelper.Duplicate(application.ActivePage, dest_pagename);
            }
        }

        public void Duplicate(string dest_pagename, IVisio.Document dest_doc)
        {
            if (!this.Session.HasActiveDrawing)
            {
                return;
            }

            if (dest_doc==null)
            {
                throw new System.ArgumentNullException("dest_doc");
            }

            var application = this.Session.VisioApplication;

            if (application.ActiveDocument == dest_doc)
            {
                throw new VA.AutomationException("dest doc is same as active doc");
            }

            var src_page = application.ActivePage;

            dest_pagename = dest_pagename ?? src_page.NameU;
            var dest_pages = dest_doc.Pages;
            var dest_page = dest_pages[1];
            VA.Pages.PageHelper.DuplicateToDocument(src_page, dest_page, dest_pagename, true);

            // the active window will be to the new document
            var active_window = application.ActiveWindow;
            active_window.Page = dest_page;
        }

        public VA.Pages.PrintPageOrientation GetOrientation()
        {
            if (!this.Session.HasActiveDrawing)
            {
                throw new AutomationException("No active page");
            }

            var application = this.Session.VisioApplication;
            var active_page = application.ActivePage;
            return VA.Pages.PageHelper.GetOrientation(active_page);
        }

        public void SetOrientation(VA.Pages.PrintPageOrientation orientation)
        {
            if (!this.Session.HasActiveDrawing)
            {
                return;
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope("Set Page Orientation"))
            {
                var active_page = application.ActivePage;
                VA.Pages.PageHelper.SetOrientation(active_page, orientation);
            }
        }

        public void ResizeToFitContents(VA.Drawing.Size bordersize, bool zoom_to_page)
        {
            if (!this.Session.HasActiveDrawing)
            {
                return;
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope("Resize Page to Fit Contents"))
            {
                var active_page = application.ActivePage;
                active_page.ResizeToFitContents(bordersize);
                if (zoom_to_page)
                {
                    this.Session.View.Zoom(VA.Scripting.Zoom.ToPage);
                }
            }
        }

        public void ResetOrigin()
        {
            if (!this.Session.HasActiveDrawing)
            {
                return;
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope("Reset Page Origin"))
            {
                var active_page = application.ActivePage;
                VA.Pages.PageHelper.ResetOrigin(active_page);
            }
        }

        public void SetSize(VA.Drawing.Size new_size)
        {
            if (!this.Session.HasActiveDrawing)
            {
                return;
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope("Set Page Size"))
            {
                var active_page = application.ActivePage;
                VA.Pages.PageHelper.SetSize(active_page,new_size);
            }
        }

        public void SetSize(double? width, double? height)
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
            var old_size = VA.Pages.PageHelper.GetSize(active_page);
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

        public void GoTo(Pages.PageNavigation flags)
        {
            var application = this.Session.VisioApplication;
            var active_document = application.ActiveDocument;
            var docpages = active_document.Pages;
            if (docpages.Count < 2)
            {
                return;
            }

            var pages = docpages;
            VA.Pages.PageHelper.NavigateTo(pages, flags);
        }
    }
}
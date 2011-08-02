using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class ApplicationCommands : CommandSet
    {
        public ApplicationCommands(Session session) :
            base(session)
        {

        }

        public string GetWindowText()
        {
            return VA.ApplicationHelper.GetApplicationWindowText(this.Session.VisioApplication);
        }

        public void WindowToFront()
        {
            var app = this.Session.VisioApplication;

            if (app == null)
            {
                return;
            }

            VA.UI.UserInterfaceHelper.BringApplicationWindowToFront(app);
        }

        public void ForceClose()
        {
            var application = this.Session.VisioApplication;
            var documents = application.Documents;
            VA.DocumentHelper.ForceCloseAll(documents);
            application.Quit(true);
            this.Session.VisioApplication = null;
        }

        public System.Drawing.Size GetWindowSize()
        {
            var rect = this.Session.VisioApplication.Window.GetWindowRect();
            var size = new System.Drawing.Size(rect.Width, rect.Height);
            return size;
        }

        /// <summary>
        /// Sets the width and height (in pixels) of the attached Visio application window
        /// </summary>
        /// <param name="scripting_session"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public void SetWindowSize(int width, int height)
        {
            if (width <= 0)
            {
                this.Session.Write(OutputStream.Error, "width must be positive");
                return;
            }

            if (height <= 0)
            {
                this.Session.Write(OutputStream.Error, "height must be positive");
                return;
            }

            var r = this.Session.VisioApplication.Window.GetWindowRect();
            r.Width = width;
            r.Height = height;
            this.Session.VisioApplication.Window.SetWindowRect(r);
        }

        public IVisio.Application AttachToRunningApplication()
        {
            var app = ApplicationHelper.FindRunningApplication();
            if (app == null)
            {
                throw new AutomationException("Did not find a running instance of Visio 2007");
            }

            this.Session.VisioApplication = app;

            VA.UI.UserInterfaceHelper.BringApplicationWindowToFront(app);

            return app;
        }

        public IVisio.Application NewApplication()
        {
            var app = new IVisio.ApplicationClass();
            this.Session.VisioApplication = app;
            return app;
        }

        public void Undo()
        {
            this.Session.VisioApplication.Undo();
        }

        public void Redo()
        {
            this.Session.VisioApplication.Redo();
        }

    }
}
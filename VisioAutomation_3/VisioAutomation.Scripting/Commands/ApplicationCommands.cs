using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VAS = VisioAutomation.Scripting;

namespace VisioAutomation.Scripting.Commands
{
    public class ApplicationCommands : SessionCommands
    {
        public ApplicationCommands(Session session) :
            base(session)
        {

        }

        public string GetApplicationWindowText()
        {
            return VA.ApplicationHelper.GetApplicationWindowText(this.Session.Application);
        }

        public void BringApplicationWindowToFront()
        {
            var app = this.Session.Application;

            if (app == null)
            {
                return;
            }

            VA.UI.UserInterfaceHelper.BringApplicationWindowToFront(app);
        }

        public void ForceApplicationClose()
        {
            var application = this.Session.Application;
            var documents = application.Documents;
            VA.DocumentHelper.ForceCloseAll(documents);
            application.Quit(true);
            this.Session.Application = null;
        }

        public System.Drawing.Size GetApplicationWindowSize()
        {
            var rect = this.Session.Application.Window.GetWindowRect();
            var size = new System.Drawing.Size(rect.Width, rect.Height);
            return size;
        }

        /// <summary>
        /// Sets the width and height (in pixels) of the attached Visio application window
        /// </summary>
        /// <param name="scripting_session"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public void SetApplicationWindowSize(int width, int height)
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

            var r = this.Session.Application.Window.GetWindowRect();
            r.Width = width;
            r.Height = height;
            this.Session.Application.Window.SetWindowRect(r);
        }

        public IVisio.Application AttachToRunningApplication()
        {
            var app = ApplicationHelper.FindRunningApplication();
            if (app == null)
            {
                throw new AutomationException("Did not find a running instance of Visio 2007");
            }

            this.Session.Application = app;

            VA.UI.UserInterfaceHelper.BringApplicationWindowToFront(app);

            return app;
        }

        public IVisio.Application StartNewApplication()
        {
            var app = new IVisio.ApplicationClass();
            this.Session.Application = app;
            return app;
        }

    }
}
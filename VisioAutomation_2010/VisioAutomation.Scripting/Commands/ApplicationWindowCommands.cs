using VisioAutomation.Application;
using VisioAutomation.Extensions;

namespace VisioAutomation.Scripting.Commands
{
    public class ApplicationWindowCommands : CommandSet
    {
        public ApplicationWindowCommands(Session session) :
            base(session)
        {

        }

        public void ToFront()
        {
            var app = this.Session.VisioApplication;

            if (app == null)
            {
                return;
            }

            ApplicationHelper.BringWindowToTop(app);
        }

        public System.Drawing.Size GetSize()
        {
            var app = this.Session.VisioApplication;
            var appwindow = app.Window;
            var rect = appwindow.GetWindowRect();
            var size = new System.Drawing.Size(rect.Width, rect.Height);
            return size;
        }

        public void SetSize(int width, int height)
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

            var app = this.Session.VisioApplication;
            var appwindow = app.Window;
            var r = appwindow.GetWindowRect();
            r.Width = width;
            r.Height = height;
            appwindow.SetWindowRect(r);
        }

    }
}
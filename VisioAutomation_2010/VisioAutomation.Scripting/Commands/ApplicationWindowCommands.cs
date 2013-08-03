using VA=VisioAutomation;
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
            this.CheckVisioApplicationAvailable();
            var app = this.Session.VisioApplication;

            if (app == null)
            {
                return;
            }

            VA.Application.ApplicationHelper.BringWindowToTop(app);
        }

        public System.Drawing.Size GetSize()
        {
            this.CheckVisioApplicationAvailable();
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
                this.Session.WriteError( "width must be positive");
                return;
            }

            if (height <= 0)
            {
                this.Session.WriteError("height must be positive");
                return;
            }

            this.CheckVisioApplicationAvailable();

            var app = this.Session.VisioApplication;
            var appwindow = app.Window;
            var r = appwindow.GetWindowRect();
            r.Width = width;
            r.Height = height;
            appwindow.SetWindowRect(r);
        }

    }
}
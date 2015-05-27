using VisioAutomation.Extensions;

namespace VisioAutomation.Scripting.Commands
{
    public class ApplicationWindowCommands : CommandSet
    {
        public ApplicationWindowCommands(Client client) :
            base(client)
        {

        }

        public void ToFront()
        {
            this.Client.Application.AssertApplicationAvailable();
            var app = this.Client.Application.Get();

            if (app == null)
            {
                return;
            }

            Application.ApplicationHelper.BringWindowToTop(app);
        }

        public System.Drawing.Size GetSize()
        {
            this.Client.Application.AssertApplicationAvailable();
            var app = this.Client.Application.Get();
            var appwindow = app.Window;
            var rect = appwindow.GetWindowRect();
            var size = new System.Drawing.Size(rect.Width, rect.Height);
            return size;
        }

        public void SetSize(int width, int height)
        {
            if (width <= 0)
            {
                this.Client.WriteError( "width must be positive");
                return;
            }

            if (height <= 0)
            {
                this.Client.WriteError("height must be positive");
                return;
            }

            this.Client.Application.AssertApplicationAvailable();

            var app = this.Client.Application.Get();
            var appwindow = app.Window;
            var r = appwindow.GetWindowRect();
            r.Width = width;
            r.Height = height;
            appwindow.SetWindowRect(r);
        }

    }
}
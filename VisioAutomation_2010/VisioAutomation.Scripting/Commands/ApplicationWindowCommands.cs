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
            this._client.Application.AssertApplicationAvailable();
            var app = this._client.Application.Get();

            if (app == null)
            {
                return;
            }

            Application.ApplicationHelper.BringWindowToTop(app);
        }

        public System.Drawing.Size GetSize()
        {
            this._client.Application.AssertApplicationAvailable();
            var app = this._client.Application.Get();
            var appwindow = app.Window;
            var rect = appwindow.GetWindowRect();
            var size = new System.Drawing.Size(rect.Width, rect.Height);
            return size;
        }

        public void SetSize(int width, int height)
        {
            if (width <= 0)
            {
                this._client.WriteError( "width must be positive");
                return;
            }

            if (height <= 0)
            {
                this._client.WriteError("height must be positive");
                return;
            }

            this._client.Application.AssertApplicationAvailable();

            var app = this._client.Application.Get();
            var appwindow = app.Window;
            var r = appwindow.GetWindowRect();
            r.Width = width;
            r.Height = height;
            appwindow.SetWindowRect(r);
        }

    }
}
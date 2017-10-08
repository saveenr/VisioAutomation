using VisioAutomation.Extensions;

namespace VisioScripting.Commands
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

            VisioAutomation.Application.ApplicationHelper.BringWindowToTop(app);
        }

        public System.Drawing.Rectangle GetRectangle()
        {
            this._client.Application.AssertApplicationAvailable();
            var app = this._client.Application.Get();
            var appwindow = app.Window;
            var rect = appwindow.GetWindowRect();
            return rect;
        }

        public void SetRectangle(System.Drawing.Rectangle rect)
        {
            this._client.Application.AssertApplicationAvailable();

            var app = this._client.Application.Get();
            var appwindow = app.Window;
            appwindow.SetWindowRect(rect);
        }
    }
}
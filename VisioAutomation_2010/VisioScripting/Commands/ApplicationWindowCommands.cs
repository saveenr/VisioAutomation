using VisioAutomation.Extensions;

namespace VisioScripting.Commands
{
    public class ApplicationWindowCommands : CommandSet
    {
        public ApplicationWindowCommands(Client client) :
            base(client)
        {

        }

        public void MoveApplicationWindowToFront()
        {
            var cmdtarget = this._client.GetCommandTargetApplication();

            var app = cmdtarget.Application;

            if (app == null)
            {
                return;
            }

            VisioAutomation.Application.ApplicationHelper.BringWindowToTop(app);
        }

        public System.Drawing.Rectangle GetApplicationWindowRectangle()
        {
            var cmdtarget = this._client.GetCommandTargetApplication();

            var appwindow = cmdtarget.Application.Window;
            var rect = appwindow.GetWindowRect();
            return rect;
        }

        public void SetApplicationWindowRectangle(System.Drawing.Rectangle rect)
        {
            var cmdtarget = this._client.GetCommandTargetApplication();

            var appwindow = cmdtarget.Application.Window;
            appwindow.SetWindowRect(rect);
        }
    }
}
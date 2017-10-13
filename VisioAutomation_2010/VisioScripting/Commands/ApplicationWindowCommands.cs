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
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application);

            var app = cmdtarget.Application;

            if (app == null)
            {
                return;
            }

            VisioAutomation.Application.ApplicationHelper.BringWindowToTop(app);
        }

        public System.Drawing.Rectangle GetRectangle()
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application);

            var appwindow = cmdtarget.Application.Window;
            var rect = appwindow.GetWindowRect();
            return rect;
        }

        public void SetRectangle(System.Drawing.Rectangle rect)
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application);

            var appwindow = cmdtarget.Application.Window;
            appwindow.SetWindowRect(rect);
        }
    }
}
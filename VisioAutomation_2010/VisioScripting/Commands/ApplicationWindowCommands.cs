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
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application);

            var app = cmdtarget.Application;

            if (app == null)
            {
                return;
            }

            VisioAutomation.Application.ApplicationHelper.BringWindowToTop(app);
        }

        public System.Drawing.Rectangle GetRectangle()
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application);

            var appwindow = cmdtarget.Application.Window;
            var rect = appwindow.GetWindowRect();
            return rect;
        }

        public void SetRectangle(System.Drawing.Rectangle rect)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application);


            var app = cmdtarget.Application;
            var appwindow = app.Window;
            appwindow.SetWindowRect(rect);
        }
    }
}
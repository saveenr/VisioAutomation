using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class GroupingCommands: CommandSet
    {
        internal GroupingCommands(Client client) :
            base(client)
        {

        }


        public IVisio.Shape Group()
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);

            // No shapes provided, use the active selection
            if (!this._client.Selection.HasShapes())
            {
                throw new VisioAutomation.Exceptions.VisioOperationException("No Selected Shapes to Group");
            }

            // the other way of doing this: this.Client.VisioApplication.DoCmd((short)IVisio.VisUICmds.visCmdObjectGroup);
            // but it doesn't return the group

            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            var g = selection.Group();
            return g;
        }

        public void Ungroup(VisioScripting.Models.TargetShapes targets)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application);

            if (targets.Shapes == null)
            {
                if (this._client.Selection.HasShapes())
                {
                    var application = cmdtarget.Application;
                    application.DoCmd((short)IVisio.VisUICmds.visCmdObjectUngroup);
                }
            }
            else
            {
                foreach (var shape in targets.Shapes)
                {
                    shape.Ungroup();
                }
            }
        }
    }
}
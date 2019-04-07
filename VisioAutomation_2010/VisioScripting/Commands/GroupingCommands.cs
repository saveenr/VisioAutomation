using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class GroupingCommands: CommandSet
    {
        internal GroupingCommands(Client client) :
            base(client)
        {

        }


        public IVisio.Shape GroupShapes(TargetSelection targetselection)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();

            // No shapes provided, use the active selection

            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            if (selection.Count<1)
            {
                throw new VisioAutomation.Exceptions.VisioOperationException("No Selected Shapes to Group");
            }

            // the other way of doing this: this.Client.VisioApplication.DoCmd((short)IVisio.VisUICmds.visCmdObjectGroup);
            // but it doesn't return the group

            var g = selection.Group();
            return g;
        }

        public void Ungroup(TargetShapes targetshapes)
        {
            targetshapes = targetshapes.Resolve(this._client);
            foreach (var shape in targetshapes.Shapes)
            {
                shape.Ungroup();
            }
        }

        public void Ungroup(TargetSelection target_selection)
        {
            var cmdtarget = this._client.GetCommandTargetApplication();
            var application = cmdtarget.Application;
            application.DoCmd((short)IVisio.VisUICmds.visCmdObjectUngroup);
        }
    }
}
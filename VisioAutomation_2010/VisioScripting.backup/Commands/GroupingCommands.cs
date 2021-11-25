using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class GroupingCommands: CommandSet
    {
        internal GroupingCommands(Client client) :
            base(client)
        {

        }


        public IVisio.Shape Group(TargetSelection targetselection)
        {
            targetselection = targetselection.ResolveToSelection(this._client);

            // the other way of doing this: this.Client.VisioApplication.DoCmd((short)IVisio.VisUICmds.visCmdObjectGroup);
            // but it doesn't return the group

            var g = targetselection.Selection.Group();
            return g;
        }

        public void Ungroup(TargetShapes targetshapes)
        {
            targetshapes = targetshapes.ResolveToShapes(this._client);
            foreach (var shape in targetshapes.Shapes)
            {
                shape.Ungroup();
            }
        }

        public void Ungroup(TargetSelection target_selection)
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.RequireApplication);
            var application = cmdtarget.Application;
            application.DoCmd((short)IVisio.VisUICmds.visCmdObjectUngroup);
        }
    }
}
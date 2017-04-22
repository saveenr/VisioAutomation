using VisioScripting.Models;
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
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            // No shapes provided, use the active selection
            if (!this._client.Selection.HasShapes())
            {
                throw new VisioAutomation.Exceptions.VisioOperationException("No Selected Shapes to Group");
            }

            // the other way of doing this: this.Client.VisioApplication.DoCmd((short)IVisio.VisUICmds.visCmdObjectGroup);
            // but it doesn't return the group

            var selection = this._client.Selection.Get();
            var g = selection.Group();
            return g;
        }

        public void Ungroup(TargetShapes targets)
        {
            this._client.Application.AssertApplicationAvailable();
            if (targets.Shapes == null)
            {
                if (this._client.Selection.HasShapes())
                {
                    var application = this._client.Application.Get();
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
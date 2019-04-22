using VisioAutomation.Application;
using VA = VisioAutomation;

namespace VisioScripting.Commands
{
    public class UndoCommands : CommandSet
    {

        public UndoCommands(Client client) :
            base(client)
        {

        }

        public void UndoLastAction()
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.RequireApplication);
            cmdtarget.Application.Undo();
        }

        public void RedoLastAction()
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.RequireApplication);
            cmdtarget.Application.Redo();
        }

        public UndoScope NewUndoScope(string name)
        {
            var app = this._client.Application.GetAttachedApplication();
            if (app == null)
            {
                throw new System.ArgumentException("Cant create UndoScope. There is no visio application attached.");
            }

            return new VA.Application.UndoScope(app, name);
        }
    }
}
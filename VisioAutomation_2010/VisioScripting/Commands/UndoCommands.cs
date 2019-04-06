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
            var cmdtarget = this._client.GetCommandTargetApplication();
            cmdtarget.Application.Undo();
        }

        public void RedoLastAction()
        {
            var cmdtarget = this._client.GetCommandTargetApplication();
            cmdtarget.Application.Redo();
        }

        public VA.Application.UndoScope NewUndoScope(string name)
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
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
    }
}
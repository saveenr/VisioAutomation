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
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application);
            cmdtarget.Application.Undo();
        }

        public void RedoLastAction()
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application);
            cmdtarget.Application.Redo();
        }
    }
}
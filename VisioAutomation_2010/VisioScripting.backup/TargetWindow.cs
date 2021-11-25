using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting
{
    public class TargetWindow : TargetObject<IVisio.Window>
    {
        private TargetWindow()
        {
        }

        internal TargetWindow(IVisio.Window window) : base(window)
        {
        }

        public TargetWindow ResolveToWindow(VisioScripting.Client client)
        {
            if (this.Resolved)
            {
                return this;
            }

            var cmdtarget = client.GetCommandTarget(CommandTargetFlags.RequireDocument);
            var active_window = cmdtarget.Application.ActiveWindow;

            client.Output.WriteVerbose("Resolving to active window (caption=\"{0}\")", active_window.Caption);

            return new TargetWindow(active_window);
        }

        public IVisio.Window Window => this._get_item_safe();

        public static TargetWindow Auto => new TargetWindow();
    }
}
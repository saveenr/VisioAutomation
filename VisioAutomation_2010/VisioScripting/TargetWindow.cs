using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting
{
    public class TargetWindow
    {
        private IVisio.Window _window;
        public TargetWindow()
        {
        }
        internal TargetWindow(IVisio.Window window)
        {
            this._window = window;
        }

        public TargetWindow Resolve(VisioScripting.Client client)
        {
            if (this._window != null)
            {
                return this;
            }

            var cmdtarget = client.GetCommandTargetDocument();
            var active_window = cmdtarget.Application.ActiveWindow;

            client.Output.WriteVerbose("Resolving to active window (caption=\"{0}\")", active_window.Caption);

            return new TargetWindow(active_window);
        }

        public IVisio.Window Window => this._window;
    }
}
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting
{
    public class TargetActiveWindow
    {
        private IVisio.Window _window;
        public TargetActiveWindow()
        {
        }
        internal TargetActiveWindow(IVisio.Window window)
        {
            this._window = window;
        }

        public TargetActiveWindow Resolve(VisioScripting.Client client)
        {
            if (this._window != null)
            {
                return this;
            }

            var cmdtarget = client.GetCommandTargetDocument();
            var active_window = cmdtarget.Application.ActiveWindow;
            return new TargetActiveWindow(active_window);
        }

        public IVisio.Window Window => this._window;
    }
}
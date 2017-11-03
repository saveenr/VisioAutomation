using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioScripting.Commands
{
    public class ViewCommands : CommandSet
    {
        private readonly double ZoomIncrement;

        internal ViewCommands(Client client) :
            base(client)
        {
        }

        public IVisio.Window GetActiveWindow()
        {
            var cmdtarget = this._client.GetCommandTargetApplication();
            var active_window = cmdtarget.Application.ActiveWindow;
            return active_window;
        }

        private static void SetViewRectToSelection(
            IVisio.Window window,
            IVisio.VisBoundingBoxArgs bbargs, 
            double padding_scale)
        {
            if (padding_scale < 0.0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(padding_scale));
            }

            if (padding_scale > 1.0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(padding_scale));
            }

            var app = window.Application;
            var active_window = app.ActiveWindow;
            var sel = active_window.Selection;
            var sel_bb = sel.GetBoundingBox(bbargs);

            var delta = sel_bb.Size * (new VisioAutomation.Geometry.Size(padding_scale,padding_scale));
            var view_rect = new VisioAutomation.Geometry.Rectangle(sel_bb.Left - delta.Width, sel_bb.Bottom - delta.Height,
                                                          sel_bb.Right + delta.Height, sel_bb.Top + delta.Height);
            window.SetViewRect(view_rect);
        }

        public void SetActiveWindowToZoom(double amount)
        {
            if (amount <= 0)
            {
                throw new System.ArgumentException("Must have positive zoom");
            }

            var cmdtarget = this._client.GetCommandTargetDocument();
            var active_window = cmdtarget.Application.ActiveWindow;
            active_window.Zoom = amount;
        }

        public double GetActiveWindowZoom()
        {
            var cmdtarget = this._client.GetCommandTargetDocument();
            var active_window = cmdtarget.Application.ActiveWindow;
            return active_window.Zoom;
        }

        public void ZoomActiveWindowRelative(double scale)
        {
            if (scale <= 0)
            {
                throw new System.ArgumentException("Must have positive scale");
            }
            var cmdtarget = this._client.GetCommandTargetDocument();
            var active_window = cmdtarget.Application.ActiveWindow;
            double old_zoom = active_window.Zoom;
            double new_zoom = old_zoom * scale;
            active_window.Zoom = new_zoom;
        }

        public void ZoomActiveWindowToObject(Models.Zoom zoom)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();
            var active_window = cmdtarget.Application.ActiveWindow;

            if (zoom == Models.Zoom.ToPage)
            {
                active_window.ViewFit = (short)IVisio.VisWindowFit.visFitPage;
            }
            else if (zoom == Models.Zoom.ToWidth)
            {
                active_window.ViewFit = (short)IVisio.VisWindowFit.visFitWidth;
            }
            else if (zoom == Models.Zoom.ToSelection)
            {
                var window = cmdtarget.Application.ActiveWindow;
                var selection = window.Selection;
                if (selection.Count<1)
                {
                    return;
                }

                double padding_scale = 0.1;
                ViewCommands.SetViewRectToSelection(active_window, IVisio.VisBoundingBoxArgs.visBBoxExtents, padding_scale);

            }
            else
            {
                throw new System.ArgumentOutOfRangeException(nameof(zoom));
            }            
        }
    }
}
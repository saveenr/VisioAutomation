using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Scripting.Commands
{
    public class ViewCommands : CommandSet
    {
        private readonly double ZoomIncrement;

        internal ViewCommands(Client client) :
            base(client)
        {
            this.ZoomIncrement = 1.20;
        }

        public IVisio.Window GetActiveWindow()
        {
            this.Client.Application.AssertApplicationAvailable();

            var application = this.Client.Application.Get();
            var active_window = application.ActiveWindow;
            return active_window;
        }

        public double GetActiveZoom()
        {
            this.Client.Application.AssertApplicationAvailable();

            var active_window = this.GetActiveWindow();
            return active_window.Zoom;
        }

        private static void SetViewRectToSelection(IVisio.Window window,
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

            var delta = sel_bb.Size*padding_scale;
            var view_rect = new Drawing.Rectangle(sel_bb.Left - delta.Width, sel_bb.Bottom - delta.Height,
                                                          sel_bb.Right + delta.Height, sel_bb.Top + delta.Height);
            window.SetViewRect(view_rect);
        }

        public void ZoomToPercentage(double amount)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (amount <= 0)
            {
                return;
            }

            var active_window = this.GetActiveWindow();
            active_window.Zoom = amount;
        }

        public void Zoom(Zoom zoom)
        {
            this.Client.Application.AssertApplicationAvailable();

            var active_window = this.GetActiveWindow();

            if (zoom == Scripting.Zoom.Out)
            {
                var cur = active_window.Zoom;
                this.ZoomToPercentage(cur / this.ZoomIncrement);                
            }
            else if (zoom == Scripting.Zoom.In)
            {
                var cur = active_window.Zoom;
                this.ZoomToPercentage(cur * this.ZoomIncrement);
            }
            else if (zoom == Scripting.Zoom.ToPage)
            {
                active_window.ViewFit = (short)IVisio.VisWindowFit.visFitPage;
            }
            else if (zoom == Scripting.Zoom.ToWidth)
            {
                active_window.ViewFit = (short)IVisio.VisWindowFit.visFitWidth;
            }
            else if (zoom == Scripting.Zoom.ToSelection)
            {
                if (!this.Client.Selection.HasShapes())
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
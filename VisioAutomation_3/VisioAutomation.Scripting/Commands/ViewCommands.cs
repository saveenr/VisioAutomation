using System;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using VA =VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class ViewCommands : SessionCommands
    {
        public ViewCommands(Session session) :
            base(session)
        {
            
        }

        public IVisio.Window GetActiveWindow()
        {
            var application = this.Session.Application;
            var active_window = application.ActiveWindow;
            return active_window;
        }

        public double GetActiveZoom()
        {
            if (!this.Session.HasActiveDrawing())
            {
                throw new AutomationException("Has no active drawing");
            }

            var active_window = GetActiveWindow();
            return active_window.Zoom;
        }

        private static void SetViewRectToSelection(IVisio.Window window,
                                                   IVisio.VisBoundingBoxArgs bbargs, double padding_scale)
        {
            if (padding_scale < 0.0)
            {
                throw new ArgumentOutOfRangeException("padding_scale");
            }

            if (padding_scale > 1.0)
            {
                throw new ArgumentOutOfRangeException("padding_scale");
            }

            var app = window.Application;
            var active_window = app.ActiveWindow;
            var sel = active_window.Selection;
            var sel_bb = sel.GetBoundingBox(bbargs);

            var delta = sel_bb.Size*padding_scale;
            var view_rect = new VA.Drawing.Rectangle(sel_bb.Left - delta.Width, sel_bb.Bottom - delta.Height,
                                                          sel_bb.Right + delta.Height, sel_bb.Top + delta.Height);
            window.SetViewRect(view_rect);
        }

        public void ZoomToPercentage(double amount)
        {
            if (!this.Session.HasActiveDrawing())
            {
                return;   
            }

            if (amount <= 0)
            {
                return;
            }

            var active_window = GetActiveWindow();
            active_window.Zoom = amount;
        }

        public void Zoom(Zoom zoom)
        {
            if (!this.Session.HasActiveDrawing())
            {
                return;
            }

            if (zoom == Scripting.Zoom.Out)
            {
                var active_window = GetActiveWindow();
                var cur = active_window.Zoom;
                ZoomToPercentage(cur / 1.20);                
            }
            else if (zoom == Scripting.Zoom.In)
            {
                var active_window = GetActiveWindow();
                var cur = active_window.Zoom;
                ZoomToPercentage(cur * 1.20);
            }
            else if (zoom == Scripting.Zoom.ToPage)
            {
                var active_window = GetActiveWindow();
                active_window.ViewFit = (short)IVisio.VisWindowFit.visFitPage;
            }
            else if (zoom == Scripting.Zoom.ToWidth)
            {
                var active_window = GetActiveWindow();
                active_window.ViewFit = (short)IVisio.VisWindowFit.visFitWidth;
            }
            else if (zoom == Scripting.Zoom.ToSelection)
            {
                if (!this.Session.HasSelectedShapes())
                {
                    return;
                }

                var window = GetActiveWindow();
                double padding_scale = 0.1;
                SetViewRectToSelection(window, IVisio.VisBoundingBoxArgs.visBBoxExtents, padding_scale);

            }
            else
            {
                throw new ArgumentOutOfRangeException("zoom");
            }            
        }
    }
}
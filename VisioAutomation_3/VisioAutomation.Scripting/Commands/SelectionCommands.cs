using System;
using System.Collections.Generic;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Linq;

namespace VisioAutomation.Scripting.Commands
{
    public class SelectionCommands : SessionCommands
    {
        public SelectionCommands(Session session) :
            base(session)
        {

        }
        
        public IEnumerable<IVisio.Shape> EnumSelectedShapes()
        {
            var app = this.Application;
            var activewin = app.ActiveWindow;
            var sel = activewin.Selection;

            var shapes = sel.AsEnumerable();
            return shapes;
        }

        public IEnumerable<IVisio.Shape> EnumSelectedShapes2D()
        {
            var shapes = this.EnumSelectedShapes().Where(s => s.OneD == 0);
            return shapes;
        }

        public IVisio.Selection GetSelection()
        {
            var application = this.Application;
            var active_window = application.ActiveWindow;
            var selection = active_window.Selection;
            return selection;
        }

        public void SelectAll()
        {
            if (!HasActiveDrawing())
            {
                return;
            }

            var active_window = this.Session.View.GetActiveWindow();
            active_window.SelectAll();
        }

        public void SelectInvert()
        {
            if (!HasActiveDrawing())
            {
                return;
            }

            var application = Application;
            var active_page = application.ActivePage;
            var shapes = active_page.Shapes;
            if (shapes.Count < 1)
            {
                return;
            }

            InvertSelection(application.ActiveWindow);
        }

        public static void InvertSelection(IVisio.Window window)
        {
            if (window == null)
            {
                throw new System.ArgumentNullException("window");
            }

            if (window.Page == null)
            {
                throw new System.ArgumentException("Window has null page", "window");
            }

            var page = (IVisio.Page) window.Page;
            var shapes = page.Shapes;
            var all_shapes = shapes.AsEnumerable();
            var selection = window.Selection;
            var selected_set = new System.Collections.Generic.HashSet<IVisio.Shape>(selection.AsEnumerable());
            var shapes_to_select = all_shapes.Where(shape => !selected_set.Contains(shape)).ToList();

            window.DeselectAll();
            window.Select(shapes_to_select, IVisio.VisSelectArgs.visSelect);
        }

        public void SelectNone()
        {
            if (!HasActiveDrawing())
            {
                return;
            }

            var application = Application;
            var active_window = application.ActiveWindow;
            active_window.DeselectAll();
            active_window.DeselectAll();
        }

        public void SelectShape(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            if (!HasActiveDrawing())
            {
                return;
            }

            var application = Application;
            var active_window = application.ActiveWindow;
            active_window.Select(shape, (short) IVisio.VisSelectArgs.visSelect);
        }

        public void SelectShapes(IEnumerable<IVisio.Shape> shapes)
        {
            if (shapes == null)
            {
                throw new ArgumentNullException("shapes");
            }

            if (!HasActiveDrawing())
            {
                return;
            }

            var application = Application;
            var active_window = application.ActiveWindow;
            active_window.Select(shapes, IVisio.VisSelectArgs.visSelect);
        }

        public void SelectShapes(IEnumerable<int> shapeids)
        {
            if (shapeids == null)
            {
                throw new ArgumentNullException("shapeids");
            }

            if (!HasActiveDrawing())
            {
                return;
            }

            var application = Application;
            var active_window = application.ActiveWindow;
            var page = application.ActivePage;
            var page_shapes = page.Shapes;
            var shapes = shapeids.Select(id => page_shapes.ItemFromID[id]).ToList();
            active_window.Select(shapes, IVisio.VisSelectArgs.visSelect);
        }
        
        public void SubSelect(IList<IVisio.Shape> shapes)
        {
            if (shapes == null)
            {
                throw new ArgumentNullException("shapes");
            }

            if (!HasActiveDrawing())
            {
                return;
            }

            Application.ActiveWindow.Select(shapes, IVisio.VisSelectArgs.visSubSelect);
        }

        public void SelectShapesByMaster(IVisio.Master master)
        {
            if (!HasActiveDrawing())
            {
                return;
            }

            var application = Application;
            var page = application.ActivePage;
            // Get a selection of connectors, by master: 
            var selection = page.CreateSelection(
                IVisio.VisSelectionTypes.visSelTypeByMaster,
                IVisio.VisSelectMode.visSelModeSkipSub, 
                master);
        }

        public void SelectShapesInLayer(string layername)
        {
            if (!HasActiveDrawing())
            {
                return;
            }

            if (layername == null)
            {
                throw new ArgumentNullException("layername");
            }

            if (layername.Length < 1)
            {
                throw new ArgumentException("layername");
            }

            var layer = this.Session.Layer.GetLayer(layername);
            var application = Application;
            var page = application.ActivePage;

            // Get a selection of connectors, by layer: 
            var selection = page.CreateSelection(
                IVisio.VisSelectionTypes.visSelTypeByLayer,
                IVisio.VisSelectMode.visSelModeSkipSub, 
                layer);
        }
    }
}
using System.Collections.Generic;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Scripting.Commands
{
    public class SelectionCommands : CommandSet
    {
        internal SelectionCommands(Client client) :
            base(client)
        {

        }
        
        public IVisio.Selection Get()
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();
            
            var application = this.Client.Application.Get();
            var active_window = application.ActiveWindow;
            var selection = active_window.Selection;
            return selection;
        }

        public void All()
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();
            
            var active_window = this.Client.View.GetActiveWindow();
            active_window.SelectAll();
        }

        public void Invert()
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.Application.Get();
            var active_page = application.ActivePage;
            var shapes = active_page.Shapes;
            if (shapes.Count < 1)
            {
                return;
            }

            SelectionCommands.Invert(application.ActiveWindow);
        }

        private static void Invert(IVisio.Window window)
        {
            if (window == null)
            {
                throw new System.ArgumentNullException(nameof(window));
            }

            if (window.Page == null)
            {
                throw new System.ArgumentException("Window has null page", nameof(window));
            }

            var page = (IVisio.Page) window.Page;
            var shapes = page.Shapes;
            var all_shapes = shapes.AsEnumerable();
            var selection = window.Selection;
            var selected_set = new HashSet<IVisio.Shape>(selection.AsEnumerable());
            var shapes_to_select = all_shapes.Where(shape => !selected_set.Contains(shape)).ToList();

            window.DeselectAll();
            window.Select(shapes_to_select, IVisio.VisSelectArgs.visSelect);
        }

        public void None()
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.Application.Get();
            var active_window = application.ActiveWindow;
            active_window.DeselectAll();
            active_window.DeselectAll();
        }

        public void Select(IVisio.Shape shape)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            var application = this.Client.Application.Get();
            var active_window = application.ActiveWindow;
            active_window.Select(shape, (short) IVisio.VisSelectArgs.visSelect);
        }

        public void Select(IEnumerable<IVisio.Shape> shapes)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (shapes == null)
            {
                throw new System.ArgumentNullException(nameof(shapes));
            }

            var application = this.Client.Application.Get();
            var active_window = application.ActiveWindow;
            active_window.Select(shapes, IVisio.VisSelectArgs.visSelect);
        }

        public void Select(IEnumerable<int> shapeids)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (shapeids == null)
            {
                throw new System.ArgumentNullException(nameof(shapeids));
            }

            var application = this.Client.Application.Get();
            var active_window = application.ActiveWindow;
            var page = application.ActivePage;
            var page_shapes = page.Shapes;
            var shapes = shapeids.Select(id => page_shapes.ItemFromID[id]).ToList();
            active_window.Select(shapes, IVisio.VisSelectArgs.visSelect);
        }
        
        public void SubSelect(IList<IVisio.Shape> shapes)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (shapes == null)
            {
                throw new System.ArgumentNullException(nameof(shapes));
            }

            var application = this.Client.Application.Get();
            var active_window = application.ActiveWindow;
            active_window.Select(shapes, IVisio.VisSelectArgs.visSubSelect);
        }

        public void SelectByMaster(IVisio.Master master)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.Application.Get();
            var page = application.ActivePage;
            // Get a selection of connectors, by master: 
            var selection = page.CreateSelection(
                IVisio.VisSelectionTypes.visSelTypeByMaster,
                IVisio.VisSelectMode.visSelModeSkipSub, 
                master);
        }

        public void SelectByLayer(string layername)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (layername == null)
            {
                throw new System.ArgumentNullException("Layer name cannot be null", "layername");
            }

            if (layername.Length < 1)
            {
                throw new System.ArgumentException("Layer name cannot be empty", nameof(layername));
            }

            var layer = this.Client.Layer.Get(layername);
            var application = this.Client.Application.Get();
            var page = application.ActivePage;

            // Get a selection of connectors, by layer: 
            var selection = page.CreateSelection(
                IVisio.VisSelectionTypes.visSelTypeByLayer,
                IVisio.VisSelectMode.visSelModeSkipSub, 
                layer);
        }

        public IList<IVisio.Shape> GetShapes()
        {
            this.Client.Application.AssertApplicationAvailable();

            var selection = this.Client.Selection.Get();
            return Selection.SelectionHelper.GetSelectedShapes(selection);
        }

        public IList<IVisio.Shape> GetShapesRecursive()
        {
            this.Client.Application.AssertApplicationAvailable();

            var selection = this.Client.Selection.Get();
            return Selection.SelectionHelper.GetSelectedShapesRecursive(selection);
        }

        public int Count()
        {
            this.Client.Application.AssertApplicationAvailable();

            var application = this.Client.Application.Get();
            var active_window = application.ActiveWindow;
            var selection = active_window.Selection;
            int count = selection.Count;
            return count;
        }

        public IList<IVisio.Shape> GetSubSelectedShapes()
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();
            
            //http://www.visguy.com/2008/05/17/detect-sub-selected-shapes-programmatically/
            var shapes = new List<IVisio.Shape>(0);
            var sel = this.Client.Selection.Get();
            var original_itermode = sel.IterationMode;

            // normal selection
            sel.IterationMode = ((short)IVisio.VisSelectMode.visSelModeSkipSub) + ((short)IVisio.VisSelectMode.visSelModeSkipSuper);

            if (sel.Count > 0)
            {
                shapes.AddRange(sel.AsEnumerable());
            }

            // sub selection
            sel.IterationMode = ((short)IVisio.VisSelectMode.visSelModeOnlySub) + ((short)IVisio.VisSelectMode.visSelModeSkipSuper);
            if (sel.Count > 0)
            {
                shapes.AddRange(sel.AsEnumerable());
            }

            sel.IterationMode = original_itermode;
            return shapes;
        }

        public void Delete()
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (!this.Client.Selection.HasShapes())
            {
                return;
            }

            var selection = this.Get();
            selection.Delete();
        }

        public void Copy()
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (!this.Client.Selection.HasShapes())
            {
                return;
            }

            var flags = IVisio.VisCutCopyPasteCodes.visCopyPasteNormal;

            var selection = this.Get();
            selection.Copy(flags);
        }

        public void Duplicate( IList<IVisio.Shape> target_shapes )
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            int n = this.GetTargetSelection(target_shapes);

            this.Client.WriteVerbose("Number of shapes to duplicate: {0}", n);

            if (n<1)
            {
                this.Client.WriteVerbose("Zero shapes to duplicate. No duplication operation performed");
                return;
            }

            var view = this.Client.View;
            var active_window = view.GetActiveWindow();
            var selection = active_window.Selection;
            selection.Duplicate();
        }

        public bool HasShapes()
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            return this.HasShapes(1);
        }

        public bool HasShapes(int min_items)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (min_items <= 0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(min_items));
            }

            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.Application.Get();
            var active_window = application.ActiveWindow;
            var selection = active_window.Selection;
            int num_selected = selection.Count;
            bool v = num_selected >= min_items;
            return v;
        }
    }
}
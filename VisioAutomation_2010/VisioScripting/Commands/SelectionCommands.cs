using System.Collections.Generic;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioScripting.Commands
{
    public class SelectionCommands : CommandSet
    {
        internal SelectionCommands(Client client) :
            base(client)
        {

        }
        
        public IVisio.Selection GetSelection(VisioScripting.TargetWindow targetwindow)
        {
            targetwindow = targetwindow.Resolve(this._client);
            var selection = targetwindow.Window.Selection;
            return selection;
        }

        public void SelectAllShapes(VisioScripting.TargetWindow targetwindow)
        {
            targetwindow = targetwindow.Resolve(this._client);

            targetwindow.Window.SelectAll();
        }

        public void InvertSelection(TargetActiveSelection taretselection)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            var active_page = cmdtarget.ActivePage;
            var shapes = active_page.Shapes;
            if (shapes.Count < 1)
            {
                return;
            }

            SelectionCommands._invert_selection(cmdtarget.Application.ActiveWindow);
        }

        private static void _invert_selection(IVisio.Window window)
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
            var all_shapes = shapes.ToEnumerable();
            var selection = window.Selection;
            var selected_set = new HashSet<IVisio.Shape>(selection.ToEnumerable());
            var shapes_to_select = all_shapes.Where(shape => !selected_set.Contains(shape)).ToList();

            window.DeselectAll();
            window.Select(shapes_to_select, IVisio.VisSelectArgs.visSelect);
        }

        public void SelectNone(VisioScripting.TargetWindow targetwindow)
        {
            targetwindow = targetwindow.Resolve(this._client);

            targetwindow.Window.DeselectAll();
            targetwindow.Window.DeselectAll();
        }

        public void SelectShapesById(VisioScripting.TargetWindow targetwindow, IVisio.Shape shape)
        {
            targetwindow = targetwindow.Resolve(this._client);


            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            targetwindow.Window.Select(shape, (short) IVisio.VisSelectArgs.visSelect);
        }

        public void SelectShapes(TargetActiveSelection targetselection, IEnumerable<IVisio.Shape> shapes)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();

            if (shapes == null)
            {
                throw new System.ArgumentNullException(nameof(shapes));
            }

            var active_window = cmdtarget.Application.ActiveWindow;
            active_window.Select(shapes, IVisio.VisSelectArgs.visSelect);
        }

        public void SelectShapesById(IEnumerable<int> shapeids)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            if (shapeids == null)
            {
                throw new System.ArgumentNullException(nameof(shapeids));
            }

            var active_window = cmdtarget.Application.ActiveWindow;
            var page = cmdtarget.ActivePage;
            var page_shapes = page.Shapes;
            var shapes = shapeids.Select(id => page_shapes.ItemFromID[id]).ToList();
            active_window.Select(shapes, IVisio.VisSelectArgs.visSelect);
        }
        
        public void SubSelectShapes(IList<IVisio.Shape> shapes)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();

            if (shapes == null)
            {
                throw new System.ArgumentNullException(nameof(shapes));
            }

            var active_window = cmdtarget.Application.ActiveWindow;
            active_window.Select(shapes, IVisio.VisSelectArgs.visSubSelect);
        }

        public void SelectShapesByMaster(TargetActiveSelection targetselection, IVisio.Master master)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            var page = cmdtarget.ActivePage;
            // Get a selection of connectors, by master: 
            var selection = page.CreateSelection(
                IVisio.VisSelectionTypes.visSelTypeByMaster,
                IVisio.VisSelectMode.visSelModeSkipSub, 
                master);
        }

        public void SelectShapesByLayer(TargetPage targetpage, string layername)
        {

            if (layername == null)
            {
                throw new System.ArgumentNullException(nameof(layername), "Layer name cannot be null" );
            }

            if (layername.Length < 1)
            {
                throw new System.ArgumentOutOfRangeException(nameof(layername), "Layer name cannot be empty");
            }

            targetpage = targetpage.Resolve(this._client);

            var layer = this._client.Layer.FindLayersOnPageByName(targetpage,layername);

            // Get a selection of connectors, by layer: 
            var selection = targetpage.Page.CreateSelection(
                IVisio.VisSelectionTypes.visSelTypeByLayer,
                IVisio.VisSelectMode.visSelModeSkipSub, 
                layer);
        }

        public IList<IVisio.Shape> GetShapes(TargetActiveSelection targetselection)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();
            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            return VisioScripting.Helpers.SelectionHelper.GetSelectedShapes(selection);
        }

        public List<IVisio.Shape> GetShapesRecursive(TargetActiveSelection targetselection)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();
            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            return VisioScripting.Helpers.SelectionHelper.GetSelectedShapesRecursive(selection);
        }

        public int GetShapeCount(TargetActiveSelection targetselection)
        {
            var cmdtarget = this._client.GetCommandTargetApplication();
            var active_window = cmdtarget.Application.ActiveWindow;
            var selection = active_window.Selection;
            int count = selection.Count;
            return count;
        }

        public List<IVisio.Shape> GetSubSelectedShapes()
        {
            var cmdtarget = this._client.GetCommandTargetDocument();

            //http://www.visguy.com/2008/05/17/detect-sub-selected-shapes-programmatically/
            var shapes = new List<IVisio.Shape>(0);
            var window = cmdtarget.Application.ActiveWindow;
            var sel = window.Selection;

            var original_itermode = sel.IterationMode;

            // normal selection
            sel.IterationMode = ((short)IVisio.VisSelectMode.visSelModeSkipSub) + ((short)IVisio.VisSelectMode.visSelModeSkipSuper);

            if (sel.Count > 0)
            {
                shapes.AddRange(sel.ToEnumerable());
            }

            // sub selection
            sel.IterationMode = ((short)IVisio.VisSelectMode.visSelModeOnlySub) + ((short)IVisio.VisSelectMode.visSelModeSkipSuper);
            if (sel.Count > 0)
            {
                shapes.AddRange(sel.ToEnumerable());
            }

            sel.IterationMode = original_itermode;
            return shapes;
        }

        public void DeleteShapes(TargetActiveSelection targetselection)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();

            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            if (selection.Count<1)
            {
                return;
            }

            selection.Delete();
        }

        public void CopySelectedShapes(TargetActiveSelection targetselection)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();

            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            if (selection.Count<1)
            {
                return;
            }

            var flags = IVisio.VisCutCopyPasteCodes.visCopyPasteNormal;
            selection.Copy(flags);
        }

        public void DuplicateShapes(TargetShapes targetshapes )
        {
            var cmdtarget = this._client.GetCommandTargetDocument();

            int n = targetshapes.SelectShapesAndCount(this._client);

            this._client.Output.WriteVerbose("Number of shapes to duplicate: {0}", n);

            if (n<1)
            {
                this._client.Output.WriteVerbose("Zero shapes to duplicate. No duplication operation performed");
                return;
            }

            var active_window = cmdtarget.Application.ActiveWindow;
            var selection = active_window.Selection;
            selection.Duplicate();
        }

        public bool ContainsShapes(TargetActiveSelection targetselection)
        {
            return this.ContainsShapes(targetselection, 1);
        }

        public bool ContainsShapes(TargetActiveSelection targetselection, int min_items)
        {
            if (min_items <= 0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(min_items));
            }

            var cmdtarget = this._client.GetCommandTargetDocument();
            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            int num_selected = selection.Count;
            bool v = num_selected >= min_items;
            return v;
        }
    }
}
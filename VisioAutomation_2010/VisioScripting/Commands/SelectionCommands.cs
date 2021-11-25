
using VisioAutomation.Extensions;

using VisioScripting.Models;

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
            targetwindow = targetwindow.ResolveToWindow(this._client);
            var selection = targetwindow.Window.Selection;
            return selection;
        }

        public void SelectShapeOperation(VisioScripting.TargetWindow targetwindow, Models.ShapeSelectionOperation operation)
        {
            if (operation == ShapeSelectionOperation.SelectAll)
            {
                this.SelectAllShapes(targetwindow);
            }
            else if (operation == ShapeSelectionOperation.InvertSelection)
            {
                this.InvertSelection(targetwindow);
            }
            else if (operation == ShapeSelectionOperation.SelectNone)
            {
                this.SelectNone(targetwindow);               
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();            }
        }

        public void SelectAllShapes(VisioScripting.TargetWindow targetwindow)
        {
            targetwindow = targetwindow.ResolveToWindow(this._client);

            targetwindow.Window.SelectAll();
        }

        public void InvertSelection(TargetWindow targetwindow)
        {
            targetwindow = targetwindow.ResolveToWindow(this._client);
            SelectionCommands._invert_selection(targetwindow.Window);
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
            targetwindow = targetwindow.ResolveToWindow(this._client);

            targetwindow.Window.DeselectAll();
            targetwindow.Window.DeselectAll();
        }

        public void SelectShapesById(VisioScripting.TargetWindow targetwindow, IVisio.Shape shape)
        {
            targetwindow = targetwindow.ResolveToWindow(this._client);


            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            targetwindow.Window.Select(shape, (short) IVisio.VisSelectArgs.visSelect);
        }

        public void SelectShapes(TargetWindow targetwindow, IEnumerable<IVisio.Shape> shapes)
        {
            if (shapes == null)
            {
                throw new System.ArgumentNullException(nameof(shapes));
            }

            targetwindow = targetwindow.ResolveToWindow(this._client);
            targetwindow.Window.Select(shapes, IVisio.VisSelectArgs.visSelect);
        }
        
        public void SubSelectShapes(IList<IVisio.Shape> shapes)
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.RequireDocument);
            if (shapes == null)
            {
                throw new System.ArgumentNullException(nameof(shapes));
            }

            var active_window = cmdtarget.Application.ActiveWindow;
            active_window.Select(shapes, IVisio.VisSelectArgs.visSubSelect);
        }

        public void SelectShapesByMaster(TargetPage targetpage, IVisio.Master master)
        {
            targetpage = targetpage.ResolveToPage(this._client);

            // Get a selection of connectors, by master: 
            var selection = targetpage.Page.CreateSelection(
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

            targetpage = targetpage.ResolveToPage(this._client);

            var layer = this._client.Layer.FindLayersOnPageByName(targetpage,layername);

            // Get a selection of connectors, by layer: 
            var selection = targetpage.Page.CreateSelection(
                IVisio.VisSelectionTypes.visSelTypeByLayer,
                IVisio.VisSelectMode.visSelModeSkipSub, 
                layer);
        }

        public IList<IVisio.Shape> GetSelectedShapes(TargetWindow targetwindow)
        {
            targetwindow = targetwindow.ResolveToWindow(this._client);

            return VisioScripting.Helpers.SelectionHelper.GetSelectedShapes(targetwindow.Window.Selection);
        }

        public IList<IVisio.Shape> GetSelectedShapes(TargetSelection targetselection)
        {
            targetselection = targetselection.ResolveToSelection(this._client);

            return VisioScripting.Helpers.SelectionHelper.GetSelectedShapes(targetselection.Selection);
        }

        public List<IVisio.Shape> GetShapesRecursive(TargetSelection targetselection)
        {
            targetselection = targetselection.ResolveToSelection(this._client);
            return VisioScripting.Helpers.SelectionHelper.GetSelectedShapesRecursive(targetselection.Selection);
        }

        public int GetShapeCount(TargetSelection targetselection)
        {
            targetselection = targetselection.ResolveToSelection(this._client);
            int count = targetselection.Selection.Count;
            return count;
        }

        public List<IVisio.Shape> GetSubSelectedShapes(TargetSelection targetselection)
        {
            targetselection = targetselection.ResolveToSelection(this._client);

            //http://www.visguy.com/2008/05/17/detect-sub-selected-shapes-programmatically/
            var shapes = new List<IVisio.Shape>(0);
            var sel = targetselection.Selection;

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

        public void DeleteShapes(TargetSelection targetselection)
        {
            targetselection = targetselection.ResolveToSelection(this._client);

            if (targetselection.Selection.Count<1)
            {
                return;
            }

            targetselection.Selection.Delete();
        }

        public void CopySelectedShapes(TargetSelection targetselection)
        {
            targetselection = targetselection.ResolveToSelection(this._client);
            if (targetselection.Selection.Count<1)
            {
                return;
            }

            var flags = IVisio.VisCutCopyPasteCodes.visCopyPasteNormal;
            targetselection.Selection.Copy(flags);
        }

        public void DuplicateShapes(TargetSelection targetselection)
        {
            targetselection = targetselection.ResolveToSelection(this._client);
            if (targetselection.Selection.Count < 1)
            {
                return;
            }

            
            this._client.Output.WriteVerbose("Number of shapes to duplicate: {0}", targetselection.Selection.Count);

            targetselection.Selection.Duplicate();
        }

        public bool ContainsShapes(TargetSelection targetselection)
        {
            return this.ContainsShapes(targetselection, 1);
        }

        public bool ContainsShapes(TargetSelection targetselection, int min_items)
        {
            if (min_items <= 0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(min_items));
            }

            targetselection = targetselection.ResolveToSelection(this._client);

            int num_selected = targetselection.Selection.Count;
            bool v = num_selected >= min_items;
            return v;
        }
    }
}
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
        
        public IVisio.Selection Get()
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            var application = cmdtarget.Application;
            var active_window = application.ActiveWindow;
            var selection = active_window.Selection;
            return selection;
        }

        public void SelectAll()
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            var active_window = this._client.View.GetActiveWindow();
            active_window.SelectAll();
        }

        public void Invert()
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            var application = cmdtarget.Application;
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
            var all_shapes = shapes.ToEnumerable();
            var selection = window.Selection;
            var selected_set = new HashSet<IVisio.Shape>(selection.ToEnumerable());
            var shapes_to_select = all_shapes.Where(shape => !selected_set.Contains(shape)).ToList();

            window.DeselectAll();
            window.Select(shapes_to_select, IVisio.VisSelectArgs.visSelect);
        }

        public void SelectNone()
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            var application = cmdtarget.Application;
            var active_window = application.ActiveWindow;
            active_window.DeselectAll();
            active_window.DeselectAll();
        }

        public void Select(IVisio.Shape shape)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            var application = cmdtarget.Application;
            var active_window = application.ActiveWindow;
            active_window.Select(shape, (short) IVisio.VisSelectArgs.visSelect);
        }

        public void Select(IEnumerable<IVisio.Shape> shapes)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            if (shapes == null)
            {
                throw new System.ArgumentNullException(nameof(shapes));
            }

            var application = cmdtarget.Application;
            var active_window = application.ActiveWindow;
            active_window.Select(shapes, IVisio.VisSelectArgs.visSelect);
        }

        public void Select(IEnumerable<int> shapeids)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            if (shapeids == null)
            {
                throw new System.ArgumentNullException(nameof(shapeids));
            }

            var application = cmdtarget.Application;
            var active_window = application.ActiveWindow;
            var page = application.ActivePage;
            var page_shapes = page.Shapes;
            var shapes = shapeids.Select(id => page_shapes.ItemFromID[id]).ToList();
            active_window.Select(shapes, IVisio.VisSelectArgs.visSelect);
        }
        
        public void SubSelect(IList<IVisio.Shape> shapes)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            if (shapes == null)
            {
                throw new System.ArgumentNullException(nameof(shapes));
            }

            var application = cmdtarget.Application;
            var active_window = application.ActiveWindow;
            active_window.Select(shapes, IVisio.VisSelectArgs.visSubSelect);
        }

        public void SelectByMaster(IVisio.Master master)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            var application = cmdtarget.Application;
            var page = application.ActivePage;
            // Get a selection of connectors, by master: 
            var selection = page.CreateSelection(
                IVisio.VisSelectionTypes.visSelTypeByMaster,
                IVisio.VisSelectMode.visSelModeSkipSub, 
                master);
        }

        public void SelectByLayer(string layername)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            if (layername == null)
            {
                throw new System.ArgumentNullException("Layer name cannot be null", nameof(layername));
            }

            if (layername.Length < 1)
            {
                throw new System.ArgumentException("Layer name cannot be empty", nameof(layername));
            }

            var layer = this._client.Layer.Get(layername);
            var application = cmdtarget.Application;
            var page = application.ActivePage;

            // Get a selection of connectors, by layer: 
            var selection = page.CreateSelection(
                IVisio.VisSelectionTypes.visSelTypeByLayer,
                IVisio.VisSelectMode.visSelModeSkipSub, 
                layer);
        }

        public IList<IVisio.Shape> GetShapes()
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);
            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            return VisioScripting.Helpers.SelectionHelper.GetSelectedShapes(selection);
        }

        public List<IVisio.Shape> GetShapesRecursive()
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);
            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            return VisioScripting.Helpers.SelectionHelper.GetSelectedShapesRecursive(selection);
        }

        public int Count()
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application);


            var application = cmdtarget.Application;
            var active_window = application.ActiveWindow;
            var selection = active_window.Selection;
            int count = selection.Count;
            return count;
        }

        public List<IVisio.Shape> GetSubSelectedShapes()
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


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

        public void Delete()
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);

            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            if (selection.Count<1)
            {
                return;
            }

            selection.Delete();
        }

        public void Copy()
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);

            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            if (selection.Count<1)
            {
                return;
            }

            var flags = IVisio.VisCutCopyPasteCodes.visCopyPasteNormal;
            selection.Copy(flags);
        }

        public void Duplicate(VisioScripting.Models.TargetShapes target_shapes )
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            int n = target_shapes.SetSelectionGetSelectedCount(this._client);

            this._client.Output.WriteVerbose("Number of shapes to duplicate: {0}", n);

            if (n<1)
            {
                this._client.Output.WriteVerbose("Zero shapes to duplicate. No duplication operation performed");
                return;
            }

            var view = this._client.View;
            var active_window = view.GetActiveWindow();
            var selection = active_window.Selection;
            selection.Duplicate();
        }

        public bool HasShapes()
        {
            return this.HasShapes(1);
        }

        public bool HasShapes(int min_items)
        {
            if (min_items <= 0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(min_items));
            }

            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);
            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            int num_selected = selection.Count;
            bool v = num_selected >= min_items;
            return v;
        }
    }
}
using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using System.Linq;

namespace VisioAutomation.Scripting
{
    public class CommandSet
    {
        // Keep a reference back to the parent client. This gives access to all other commands
        // for a the current context
        protected readonly Client Client;

        public CommandSet(Client client)
        {
            this.Client = client;
        }

        protected void AssertApplicationAvailable()
        {
            var has_app = this.Client.HasApplication;
            if (!has_app)
            {
                throw new VisioApplicationException("No Visio Application available");
            }
        }

        protected void AssertDocumentAvailable()
        {
            if (!this.Client.HasActiveDocument)
            {
                throw new VA.Scripting.ScriptingException("No Drawing available");
            }

        }

        public VA.Drawing.DrawingSurface GetDrawingSurface()
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            var surf_Application = this.Client.VisioApplication;
            var surf_Window = surf_Application.ActiveWindow;
            var surf_Window_subtype = surf_Window.SubType;
            
            // TODO: Revisit the logic here
            // TODO: And what about a selected shape as a surface?

            this.Client.WriteVerbose("Window SubType: {0}", surf_Window_subtype);
            if (surf_Window_subtype == 64)
            {
                this.Client.WriteVerbose("Window = Master Editing");
                var surf_Master = (IVisio.Master)surf_Window.Master;
                var surface = new VA.Drawing.DrawingSurface(surf_Master);
                return surface;

            }
            else
            {
                this.Client.WriteVerbose("Window = Page ");
                var surf_Page = surf_Application.ActivePage;
                var surface = new VA.Drawing.DrawingSurface(surf_Page);
                return surface;
            }
        }

        public VA.ShapeSheet.ShapeSheetSurface GetShapeSheetSurface()
        {
            var ds = this.GetDrawingSurface();
            var ss = new ShapeSheetSurface(ds.Target);
            return ss;
        }

        internal static IEnumerable<System.Reflection.MethodInfo> GetCommandMethods(System.Type mytype)
        {
            var cmdsettype = typeof(VA.Scripting.CommandSet);

            if (!cmdsettype.IsAssignableFrom(mytype))
            {
                string msg = string.Format("{0} must derive from {1}", mytype.Name, cmdsettype.Name);
            }

            var methods = mytype.GetMethods().Where(m => m.IsPublic && !m.IsStatic);

            foreach (var method in methods)
            {
                if (method.Name == "ToString" || method.Name == "GetHashCode" || method.Name == "GetType" || method.Name == "Equals")
                {
                    continue;
                }

                yield return method;
            }
        }

        protected IList<IVisio.Shape> GetTargetShapes(IList<IVisio.Shape> shapes)
        {
            this.AssertApplicationAvailable();
            if (shapes == null)
            {
                // If no collection of shapes were passed in then use the selection
                this.Client.WriteVerbose("GetTargetShapes: Targeting shapes from active selection");
                var out_shapes = this.Client.Selection.GetShapes();
                this.Client.WriteVerbose("GetTargetShapes: Number of shapes = {0}", out_shapes.Count);
                return out_shapes;
            }
            this.Client.WriteVerbose("GetTargetShapes: Targeting specified shapes ");
            this.Client.WriteVerbose("GetTargetShapes: Number of shapes = {0}", shapes.Count);
            return shapes;
        }

        protected int GetTargetSelection(IList<IVisio.Shape> shapes)
        {
            this.AssertApplicationAvailable();

            if (shapes == null)
            {
                this.Client.WriteVerbose("GetTargetSelection: Targeting shapes from active selection");
                int n = this.Client.Selection.Count();
                this.Client.WriteVerbose("GetTargetSelection: Number of shapes = {0}", n);
                return n;
            }

            this.Client.WriteVerbose("GetTargetSelection: Targeting specified shapes");
            this.Client.WriteVerbose("GetTargetSelection: Number of shapes specified = {0}", shapes.Count);
            this.Client.WriteVerbose("GetTargetSelection: Clearing Selection");
            this.Client.Selection.None();
            this.Client.WriteVerbose("GetTargetSelection: Setting selection");
            this.Client.Selection.Select(shapes);
            int n2 = this.Client.Selection.Count();
            this.Client.WriteVerbose("GetTargetSelection: Selection contains {0} shapes", n2);
            return n2;
        }

        protected IVisio.Shape GetTargetShape( IVisio.Shape shape)
        {
            this.AssertApplicationAvailable();

            if (shape == null)
            {
                this.Client.WriteVerbose("GetTargetShape: Targeting single shape from active selection");
                // If no collection of shapes were passed in then use the selection
                var out_shapes = this.Client.Selection.GetShapes();
                int n = out_shapes.Count;
                this.Client.WriteVerbose("GetTargetShape: Number of shapes from selection = {0}", n);

                if (out_shapes.Count > 0)
                {
                    this.Client.WriteVerbose("GetTargetShape: More than 1 shape in selection, targeting the first one");                    
                    return out_shapes[0];
                }

                this.Client.WriteVerbose("GetTargetShape: No shapes in selection, targeting none");
                return null;
            }
            this.Client.WriteVerbose("GetTargetShape: Targeting specified shape");
            return shape;
        }
    }
}
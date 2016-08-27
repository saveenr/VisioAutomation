using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Commands
{
    public class CommandSet
    {
        // Keep a reference back to the parent client. This gives access to all other commands
        // for a the current context
        protected readonly Client _client;

        public CommandSet(Client client)
        {
            this._client = client;
        }

        internal static IEnumerable<Command> GetCommands(System.Type mytype)
        {
            var cmdsettype = typeof(CommandSet);

            if (!cmdsettype.IsAssignableFrom(mytype))
            {
                string msg = string.Format("{0} must derive from {1}", mytype.Name, cmdsettype.Name);
            }

            var methods = mytype.GetMethods().Where(m => m.IsPublic && !m.IsStatic);

            foreach (var method in methods)
            {
                // Skip some method names
                switch (method.Name)
                {
                    case "ToString":
                    case "GetHashCode":
                    case "GetType":
                    case "Equals":
                        continue;
                }

                var cmd = new Command(method);
                yield return cmd;
            }
        }

        protected IList<IVisio.Shape> GetTargetShapes2D(IList<IVisio.Shape> shapes)
        {
            var shapes_2d = this.GetTargetShapes(shapes).Where(s => s.OneD == 0).ToList();
            return shapes_2d;
        }

        protected IList<IVisio.Shape> GetTargetShapes(IList<IVisio.Shape> shapes)
        {
            this._client.Application.AssertApplicationAvailable();

            if (shapes == null)
            {
                var out_shapes = this._client.Selection.GetShapes();
                this._client.WriteVerbose("GetTargetShapes: Returning {0} shapes from the active selection", out_shapes.Count);
                return out_shapes;
            }

            this._client.WriteVerbose("GetTargetShapes: Returning {0} shapes that were passed in", shapes.Count);
            return shapes;
        }

        protected int GetTargetSelectionCount(IList<IVisio.Shape> shapes)
        {
            this._client.Application.AssertApplicationAvailable();

            if (shapes == null)
            {
                int n = this._client.Selection.Count();
                this._client.WriteVerbose("GetTargetSelectionCount: Using active selection of {0} shapes", n);
                return n;
            }

            this._client.WriteVerbose("GetTargetSelectionCount: Reseting selecton to specified {0} shapes", shapes.Count);
            this._client.Selection.None();
            this._client.Selection.Select(shapes);
            int selected_count = this._client.Selection.Count();
            return selected_count;
        }

        protected IVisio.Shape GetTargetShape( IVisio.Shape shape)
        {
            this._client.Application.AssertApplicationAvailable();

            if (shape == null)
            {
                this._client.WriteVerbose("GetTargetShape: Targeting single shape from active selection");
                // If no collection of shapes were passed in then use the selection
                var out_shapes = this._client.Selection.GetShapes();
                int n = out_shapes.Count;
                this._client.WriteVerbose("GetTargetShape: Number of shapes from selection = {0}", n);

                if (out_shapes.Count > 0)
                {
                    this._client.WriteVerbose("GetTargetShape: More than 1 shape in selection, targeting the first one");                    
                    return out_shapes[0];
                }

                this._client.WriteVerbose("GetTargetShape: No shapes in selection, targeting none");
                return null;
            }
            this._client.WriteVerbose("GetTargetShape: Targeting specified shape");
            return shape;
        }
    }
}
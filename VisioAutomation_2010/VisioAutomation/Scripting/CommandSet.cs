using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Scripting
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
                string msg = $"{mytype.Name} must derive from {cmdsettype.Name}";
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
                // If no collection of shapes were passed in then use the selection
                this._client.WriteVerbose("GetTargetShapes: Targeting shapes from active selection");
                var out_shapes = this._client.Selection.GetShapes();
                this._client.WriteVerbose("GetTargetShapes: Number of shapes = {0}", out_shapes.Count);
                return out_shapes;
            }

            this._client.WriteVerbose("GetTargetShapes: Targeting specified shapes ");
            this._client.WriteVerbose("GetTargetShapes: Number of shapes = {0}", shapes.Count);
            return shapes;
        }

        protected int GetTargetSelection(IList<IVisio.Shape> shapes)
        {
            this._client.Application.AssertApplicationAvailable();

            if (shapes == null)
            {
                this._client.WriteVerbose("GetTargetSelection: Targeting shapes from active selection");
                int n = this._client.Selection.Count();
                this._client.WriteVerbose("GetTargetSelection: Number of shapes = {0}", n);
                return n;
            }

            this._client.WriteVerbose("GetTargetSelection: Targeting specified shapes");
            this._client.WriteVerbose("GetTargetSelection: Number of shapes specified = {0}", shapes.Count);
            this._client.WriteVerbose("GetTargetSelection: Clearing Selection");
            this._client.Selection.None();
            this._client.WriteVerbose("GetTargetSelection: Setting selection");
            this._client.Selection.Select(shapes);
            int n2 = this._client.Selection.Count();
            this._client.WriteVerbose("GetTargetSelection: Selection contains {0} shapes", n2);
            return n2;
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
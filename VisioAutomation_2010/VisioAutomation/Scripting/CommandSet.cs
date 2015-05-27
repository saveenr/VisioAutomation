using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Scripting
{
    public class Command
    {
        public System.Reflection.MethodInfo MethodInfo;

        public Command(System.Reflection.MethodInfo mi)
        {
            this.MethodInfo = mi;
        }
    }

    public class CommandSet
    {
        // Keep a reference back to the parent client. This gives access to all other commands
        // for a the current context
        protected readonly Client Client;

        public CommandSet(Client client)
        {
            this.Client = client;
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
                if (method.Name == "ToString" || method.Name == "GetHashCode" || method.Name == "GetType" || method.Name == "Equals")
                {
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
            this.Client.Application.AssertApplicationAvailable();
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
            this.Client.Application.AssertApplicationAvailable();

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
            this.Client.Application.AssertApplicationAvailable();

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
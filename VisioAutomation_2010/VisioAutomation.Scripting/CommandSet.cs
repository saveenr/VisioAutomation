using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using System.Linq;

namespace VisioAutomation.Scripting
{
    public class CommandSet
    {
        // Keep a reference back to the parent session. This gives access to all other commands
        // for a the current context
        protected readonly Session Session;

        public CommandSet(Session session)
        {
            this.Session = session;
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
            if (shapes == null)
            {
                // If no collection of shapes were passed in then use the selection
                var out_shapes = this.Session.Selection.GetShapes(VA.Selection.ShapesEnumeration.Flat);
                return out_shapes;
            }
            return shapes;
        }

        protected IVisio.Shape GetTargetShape( IVisio.Shape shape)
        {
            if (shape == null)
            {
                // If no collection of shapes were passed in then use the selection
                var out_shapes = this.Session.Selection.GetShapes(VA.Selection.ShapesEnumeration.Flat);
                if (out_shapes.Count > 0)
                {
                    
                    return out_shapes[0];
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return shape;
            }
        }
    }
}
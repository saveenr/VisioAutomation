using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using System.Linq;
namespace VisioAutomation.Scripting
{
    public class SessionCommands
    {
        // Keep a reference back to the parent session. This gives access to all other commands
        // for a the current context
        protected readonly Session Session;

        public SessionCommands(Session session)
        {
            this.Session = session;
        }       

        public virtual string GetHelp()
        {
            var lines = new List<string>();

            var mytype = this.GetType();

            // retrieve all public nonstatic methods
            var methods = mytype.GetMethods().Where(m=>m.IsPublic).Where(m=>!m.IsStatic);
            
            var sb = new System.Text.StringBuilder();
            foreach (var method in methods)
            {
                if (method.Name == "ToString" || method.Name == "GetHashCode" || method.Name == "GetType" || method.Name == "Equals")
                {
                    continue;
                }

                sb.Length = 0;
                var method_params = method.GetParameters();
                int i = 0;
                foreach (var param in method_params)
                {
                    if (i > 0)
                    {
                        sb.Append(", ");
                    }

                    string paramtext = string.Format("[{0}] {1}", param.ParameterType.Name, param.Name);
                    sb.Append(paramtext);

                    i++;
                }
                
                string line = string.Format("{0}({1})", method.Name, sb.ToString());

                lines.Add(line.ToString());

            }

            var helpstr = new System.Text.StringBuilder(lines.Select(s => s.Length + 2).Sum());
            foreach (var line in lines)
            {
                helpstr.Append(line);
                helpstr.Append("\r\n");
            }

            return helpstr.ToString();
        }
    }
}
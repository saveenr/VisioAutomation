using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class DeveloperCommands : CommandSet
    {
        public DeveloperCommands(Session session) :
            base(session)
        {

        }

        public void HelloWorld()
        {
            if (this.Session.VisioApplication == null)
            {
                this.Session.Application.New();
            }

            var doc = this.Session.Document.New(8.5, 11);
            var pages = doc.Pages;
            var page = pages.Add();

            var s0 = page.DrawRectangle(2, 2, 6, 6);
            s0.Text = "Hello World";
        }

        public System.Xml.Linq.XElement GetXMLDescription()
        {
            var el_shapes = new System.Xml.Linq.XElement("Shapes");
            if (!this.Session.HasSelectedShapes())
            {
                return el_shapes;
            }

            var page = this.Session.VisioApplication.ActivePage;
            var shapes = page.Shapes.AsEnumerable().ToList();
            var shapeids = shapes.Select(s => s.ID).ToList();

            var el_shape = VA.ShapeHelper.GetShapeDescriptionXML(page, shapeids);

            foreach (var x in el_shape)
            {
                el_shapes.Add(x);
            }

            return el_shapes;
        }


        private IEnumerable<System.Reflection.MethodInfo> get_command_methods(System.Type mytype) 
        {
            var methods = mytype.GetMethods().Where(m => m.IsPublic).Where(m => !m.IsStatic);
            foreach (var method in methods)
            {
                if (method.Name == "ToString" || method.Name == "GetHashCode" || method.Name == "GetType" || method.Name == "Equals")
                {
                    continue;
                }
                yield return method;
            }
        }

        private string get_nice_typename(System.Type t)
        {
            if (t == typeof(int))
            {
                return "int";
            }
            else if (t == typeof(string))
            {
                return "string";
            }
            else if (t == typeof(double))
            {
                return "double";
            }
            else if (t == typeof(bool))
            {
                return "bool";
            }
            else if (t == typeof(short))
            {
                return "short";
            }

            if (IsNullableType(t))
            {
                var actualtype = t.GetGenericArguments()[0];
                return get_nice_typename(actualtype)+"?";
            }

            if (t.IsGenericType)
            {
                var sb = new System.Text.StringBuilder();
                var tokens = t.Name.Split(new[] {'`'});
                
                sb.Append(tokens[0]);
                var gas = t.GetGenericArguments();
                var ga_names = gas.Select(i => get_nice_typename(i));

                sb.Append("<");
                TextUtil.Join(sb,", ",ga_names);
                sb.Append(">");
                return sb.ToString();
            }

            return t.Name;
        }

        private bool IsNullableType(System.Type colType)
        {
            if (
            (colType.IsGenericType) &&
            (colType.GetGenericTypeDefinition() == typeof(System.Nullable<>)))
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public virtual IVisio.Document Documentation()
        {
            var docbuilder = new VisioAutomation.Experimental.SimpleTextDoc.TextDocumentBuilder(this.Session.VisioApplication);
            var lines = new List<string>();

            var cmdst_props = GetCmdsetPropeties().OrderBy(i=>i.Name).ToList();
            var sb = new System.Text.StringBuilder();
            var helpstr = new System.Text.StringBuilder();

            foreach (var cmdset_prop in cmdst_props)
            {
                var cmdset_type = cmdset_prop.PropertyType;

                // Calculate the text
                var methods = this.get_command_methods(cmdset_type);
                lines.Clear();
                foreach (var method in methods)
                {
                    sb.Length = 0;
                    var method_params = method.GetParameters();
                    TextUtil.Join(sb, ", ", method_params.Select(param => string.Format("{0} {1}", get_nice_typename(param.ParameterType), param.Name)));

                    if (method.ReturnType != typeof(void))
                    {
                        string line = string.Format("{0}({1}) -> {2}", method.Name, sb, get_nice_typename(method.ReturnType));
                        lines.Add(line);
                    }
                    else
                    {
                        string line = string.Format("{0}({1})", method.Name, sb);
                        lines.Add(line);
                    }
                }

                lines.Sort();
                
                helpstr.Length = 0;
                TextUtil.Join(helpstr,"\r\n",lines);

                var xpage = new VisioAutomation.Experimental.SimpleTextDoc.TextPage();
                xpage.Title = cmdset_prop.Name + " commands";
                xpage.Body = helpstr.ToString();
                xpage.Name = cmdset_prop.Name + " commands";

                docbuilder.Draw(xpage);
            }

            docbuilder.Finish();
            docbuilder.VisioDocument.Subject = "VisioAutomation.Scripting Documenation";
            docbuilder.VisioDocument.Title = "VisioAutomation.Scripting Documenation";
            docbuilder.VisioDocument.Creator = "";
            docbuilder.VisioDocument.Company = "";

            return docbuilder.VisioDocument;
        }

        private static List<System.Reflection.PropertyInfo> GetCmdsetPropeties()
        {
            var props = typeof(VA.Scripting.Session).GetProperties()
                .Where(
                    p => typeof(VA.Scripting.CommandSet).IsAssignableFrom(p.PropertyType))
                .ToList();
            return props;
        }
    }
}

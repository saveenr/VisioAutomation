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

        public IVisio.Document DrawScriptingDocumentation()
        {
            var pagesize = new VA.Drawing.Size(8.5, 11);
            var docbuilder = new VA.Experimental.SimpleTextDoc.TextDocumentBuilder(this.Session.VisioApplication, pagesize);
            docbuilder.BodyParaSpacingAfter = 6.0;
            var lines = new List<string>();

            var cmdst_props = VA.Scripting.Session.GetCommandSetProperties().OrderBy(i=>i.Name).ToList();
            var sb = new System.Text.StringBuilder();
            var helpstr = new System.Text.StringBuilder();

            docbuilder.Start();
            foreach (var cmdset_prop in cmdst_props)
            {
                var cmdset_type = cmdset_prop.PropertyType;

                // Calculate the text
                var methods = CommandSet.GetCommandMethods(cmdset_type);
                lines.Clear();
                foreach (var method in methods)
                {
                    sb.Length = 0;
                    var method_params = method.GetParameters();
                    TextUtil.Join(sb, ", ", method_params.Select(param => string.Format("{0} {1}", ReflectionUtil.GetNiceTypeName(param.ParameterType), param.Name)));

                    if (method.ReturnType != typeof(void))
                    {
                        string line = string.Format("{0}({1}) -> {2}", method.Name, sb, ReflectionUtil.GetNiceTypeName(method.ReturnType));
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

                var docpage = new VisioAutomation.Experimental.SimpleTextDoc.TextPage();
                docpage.Title = cmdset_prop.Name + " commands";
                docpage.Body = helpstr.ToString();
                docpage.Name = cmdset_prop.Name + " commands";

                docbuilder.Draw(docpage);
            }

            docbuilder.Finish();
            docbuilder.VisioDocument.Subject = "VisioAutomation.Scripting Documenation";
            docbuilder.VisioDocument.Title = "VisioAutomation.Scripting Documenation";
            docbuilder.VisioDocument.Creator = "";
            docbuilder.VisioDocument.Company = "";

            return docbuilder.VisioDocument;
        }

        public IVisio.Document DrawInteropEnumDocumentation()
        {
            var pagesize = new VA.Drawing.Size(8.5, 11);
            var docbuilder = new VA.Experimental.SimpleTextDoc.TextDocumentBuilder(this.Session.VisioApplication, pagesize);
            //docbuilder.BodyParaSpacingAfter = 2.0;
            docbuilder.BodyTextSize = 8.0;
            var helpstr = new System.Text.StringBuilder();
            int chunksize = 70;

            var interop_enums = VA.Interop.InteropHelper.GetEnums();
            docbuilder.Start();
            int pagecount = 0;
            foreach (var enum_ in interop_enums)
            {


                int chunkcount = 0;

                var values = enum_.Values.OrderBy(i => i.Name).ToList();
                foreach (var chunk in Chunk(values, chunksize))
                {
                    helpstr.Length = 0;
                    foreach (var val in chunk)
                    {
                        helpstr.AppendFormat("0x{0}\t{1}\n", val.Value.ToString("x"),val.Name);

                    }

                    var docpage = new VA.Experimental.SimpleTextDoc.TextPage();
                    docpage.Title = enum_.Name;
                    docpage.Body = helpstr.ToString();
                    if (chunkcount == 0)
                    {
                        docpage.Name = string.Format("{0}", enum_.Name);
                        
                    }
                    else
                    {
                        docpage.Name = string.Format("{0} ({1})", enum_.Name, chunkcount + 1);
                    }

                    docbuilder.Draw(docpage);

                    var tabstops = new[]
                                 {
                                     new VA.Text.TabStop(1.5, VA.Text.TabStopAlignment.Left)
                                 };
                    VA.Text.TextHelper.SetTabStops(docpage.VisioBodyShape, tabstops);
                    
                    chunkcount++;
                    pagecount++;
                }

            }

            docbuilder.Finish();
            docbuilder.VisioDocument.Subject = "Visio Interop Enum Documenation";
            docbuilder.VisioDocument.Title = "Visio Interop Enum Documenation";
            docbuilder.VisioDocument.Creator = "";
            docbuilder.VisioDocument.Company = "";

            return docbuilder.VisioDocument;
        }

        private class PathTreeBuilder
        {
            public Dictionary<string, string> PathToParentPath;
            public List<string> Roots;
            public string Separator;
            public string[] seps;
            private System.StringSplitOptions options = System.StringSplitOptions.None;

            public PathTreeBuilder()
            {
                this.PathToParentPath = new Dictionary<string, string>();
                this.Roots = new List<string>();
                this.Separator = ".";
                this.seps = new[] {this.Separator};
            }

            public void Add(string path)
            {
                if (this.PathToParentPath.ContainsKey(path))
                {
                    return;
                }

                var tokens = path.Split(seps,options);

                if (tokens.Length == 0)
                {
                    throw new VA.AutomationException();
                }
                else if (tokens.Length == 1)
                {
                    string first = tokens[0];
                    this.Roots.Add(first);
                    this.PathToParentPath[first] = null;
                }
                else
                {
                    string parent_path = string.Join(this.Separator, tokens.Take(tokens.Length - 1));
                    this.Add(parent_path);
                    this.PathToParentPath[path] = parent_path;
                }   
            }

            public List<string> GetPaths()
            {
                return this.PathToParentPath.Keys.ToList();
            }
        }

        public IVisio.Document DrawVANamespaces()
        {
            var doc = this.Session.Document.New(8.5,11);

            var types = VA.Experimental.Developer.DeveloperHelper.GetAllTypes();
            var pathbuilder = new PathTreeBuilder();
            foreach (var type in types)
            {
                pathbuilder.Add(type.Namespace);
            }

            var namespaces = pathbuilder.GetPaths();
            
            var tree_layout = new VA.Layout.Tree.Drawing();
            tree_layout.LayoutOptions.Direction = VA.Layout.Tree.LayoutDirection.Right;
            tree_layout.LayoutOptions.UseDynamicConnectors = true;
            var ns_node_map = new Dictionary<string, VA.Layout.Tree.Node>(namespaces.Count);

            // create nodes for every namespace
            foreach (string ns in namespaces)
            {
                string label = ns;
                int index_of_last_sep = ns.LastIndexOf(pathbuilder.Separator);
                if (index_of_last_sep > 0)
                {
                    label = ns.Substring(index_of_last_sep+1);
                }

                var node = new VA.Layout.Tree.Node(ns);
                node.Text = label;
                node.Size = new VA.Drawing.Size(2.0, 0.25);
                ns_node_map[ns] = node;
            }

            // add children to nodes
            foreach (string ns in namespaces)
            {
                var parent_ns = pathbuilder.PathToParentPath[ns];

                if (parent_ns != null)
                {
                    // the current namespace has a parent
                    var parent_node = ns_node_map[parent_ns];
                    var child_node = ns_node_map[ns];
                    parent_node.Children.Add(child_node);
                }
                else
                {
                    // that means this namespace is a root, forget about it
                }
            }

            if (pathbuilder.Roots.Count == 0)
            {
                
            }
            else if (pathbuilder.Roots.Count == 1)
            {
                // when there is exactly one root namespace, then that node will be the tree's root node
                var first_root = pathbuilder.Roots[0];
                var root_n = ns_node_map[first_root];
                tree_layout.Root = root_n;
            }
            else
            {
                // if there are multiple root namespaces, inject an empty placeholder root
                var root_n = new VA.Layout.Tree.Node();
                tree_layout.Root = root_n;

                foreach (var root_ns in pathbuilder.Roots)
                {
                    var node = ns_node_map[root_ns];
                    tree_layout.Root.Children.Add(node);
                }
            }

            tree_layout.Render(doc.Application.ActivePage);
            return doc;
        }


        public IList<VA.Interop.EnumType> GetInteropEnums()
        {
            return VA.Interop.InteropHelper.GetEnums();
        }

        public VA.Interop.EnumType GetInteropEnum(string name)
        {
            return VA.Interop.InteropHelper.GetEnum(name);
        }

        private static IEnumerable<IEnumerable<T>> Chunk<T>(IEnumerable<T> source, int chunksize)
        {
            while (source.Any())
            {
                yield return source.Take(chunksize);
                source = source.Skip(chunksize);
            }
        }
    }
}


namespace VisioAutomation.Experimental.Developer
{
    public class DeveloperHelper
    {
        public static List<System.Type> GetTypes()
        {
            // find the VA assembly
            var vat = typeof (VisioAutomation.ApplicationHelper);
            var asm = vat.Assembly;

            // TODO: Consider filtering out types that should *not* be exposed despite being public
            var types = asm.GetExportedTypes().Where(t => t.IsPublic).ToList();
            return types;
        }

        public static List<System.Type> GetAllTypes()
        {
            // find the VA assembly
            var vat = typeof(VisioAutomation.ApplicationHelper);
            var asm = vat.Assembly;

            var types = asm.GetExportedTypes().ToList();
            return types;
        }

    }
}
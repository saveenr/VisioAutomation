using System;
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

        public void QuickStart()
        {
            if (this.Session.VisioApplication == null)
            {
                this.Session.Application.NewApplication();
            }

            var doc = this.Session.Document.NewDocument(8.5, 11);
            var pages = doc.Pages;
            pages.Add();
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

            return t.Name;
        }

        public virtual IVisio.Document DrawDocumentation()
        {
            var pagesize = new VA.Drawing.Size(8.5, 11);
            var pagerect = new VA.Drawing.Rectangle(new VA.Drawing.Point(0, 0), pagesize);
            var titlerect = new VA.Drawing.Rectangle(pagerect.UpperLeft.Add(0.5, -1.0), pagerect.UpperRight.Subtract(0.5, 0.5));
            var bodyrect = new VA.Drawing.Rectangle(pagerect.LowerLeft.Add(0.5, 0.5), pagerect.UpperRight.Subtract(0.5, 1.0));

            var app = this.Session.VisioApplication;
            var docs = app.Documents;
            var doc = docs.Add("");

            doc.Subject = "VisioAutomation.Scripting Documenation";
            doc.Title = "VisioAutomation.Scripting Documenation";
            doc.Creator = "";
            doc.Company = "";

            var font = doc.Fonts["Segoe UI"];
            int fontid = font.ID;

            var textblockformat = new VA.Text.TextBlockFormatCells();
            textblockformat.VerticalAlign = 0;

            var title_para_fmt = new VA.Text.ParagraphFormatCells();
            title_para_fmt.HorizontalAlign = 0;

            var title_format = new VA.Format.ShapeFormatCells();
            title_format.LineWeight = 0;
            title_format.LinePattern= 0;

            var title_char_fmt = new VA.Text.CharacterFormatCells();
            title_char_fmt.Font = fontid;
            title_char_fmt.Size = VA.Convert.PointsToInches(15.0);

            var body_para_fmt = new VA.Text.ParagraphFormatCells();
            body_para_fmt.HorizontalAlign = 0;
            body_para_fmt.SpacingAfter = VA.Convert.PointsToInches(6.0);

            var body_char_fmt = new VA.Text.CharacterFormatCells();
            body_char_fmt.Font = fontid;
            body_char_fmt.Size = VA.Convert.PointsToInches(8.0);

            var body_format = new VA.Format.ShapeFormatCells();
            body_format.LineWeight = 0;
            body_format.LinePattern = 0;

            var lines = new List<string>();

            var cmdst_props = GetCmdsetPropeties().OrderBy(i=>i.Name).ToList();
            var sb = new System.Text.StringBuilder();
            var helpstr = new System.Text.StringBuilder();


            foreach (var cmdset_prop in cmdst_props)
            {
                var cmdset_type = cmdset_prop.PropertyType;
                
                var page = doc.Pages.Add();
                page.NameU = cmdset_prop.Name + " commands";
                VA.PageHelper.SetSize(page, pagesize);

                // Calculate the text
                var methods = this.get_command_methods(cmdset_type);
                lines.Clear();
                foreach (var method in methods)
                {
                    sb.Length = 0;
                    var method_params = method.GetParameters();
                    TextUtil.Join(sb, ", ", method_params.Select(param => string.Format("{0} {1}", get_nice_typename(param.ParameterType), param.Name)));
                    string line = string.Format("{0}({1})", method.Name, sb);
                    lines.Add(line);
                }

                lines.Sort();
                
                helpstr.Length = 0;
                TextUtil.Join(helpstr,"\r\n",lines);

                // Draw the shapes
                var titleshape = page.DrawRectangle(titlerect);
                titleshape.Text = cmdset_prop.Name + " commands";

                var bodyshape = page.DrawRectangle(bodyrect);
                bodyshape.Text = helpstr.ToString();

                var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

                // Set the ShapeSheet props
                short bodyshape_id = bodyshape.ID16;
                short titleshape_id = titleshape.ID16;
                textblockformat.Apply(update, titleshape_id);
                title_para_fmt.Apply(update, titleshape_id, 0);
                title_char_fmt.Apply(update,titleshape_id,0);
                title_format.Apply(update,titleshape_id);

                textblockformat.Apply(update, bodyshape_id);
                body_para_fmt.Apply(update, bodyshape_id, 0);
                body_char_fmt.Apply(update, bodyshape_id, 0);
                body_format.Apply(update, bodyshape_id);
                update.Execute(page);
            }

            // Delete the empty first page
            var first_page = doc.Pages[1];
            first_page.Delete(1);
            first_page = null;

            // set the new first page
            first_page = doc.Pages[1];
            first_page.Activate();

            return doc;
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
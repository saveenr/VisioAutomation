using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Drawing;
using VisioAutomation.Extensions;
using VisioAutomation.Format;
using VisioAutomation.Text;
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
            var dd = new DocDoc(this.Session.VisioApplication);
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
                    string line = string.Format("{0}({1})", method.Name, sb);
                    lines.Add(line);
                }

                lines.Sort();
                
                helpstr.Length = 0;
                TextUtil.Join(helpstr,"\r\n",lines);

                var xpage = new DocPage();
                xpage.Title = cmdset_prop.Name + " commands";
                xpage.Body = helpstr.ToString();
                xpage.Name = cmdset_prop.Name + " commands";

                dd.Draw(xpage);
            }

            dd.Finish();

            return dd.doc;
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

    internal class DocPage
    {
        public string Title;
        public string Body;
        public string Name;
    }

    internal class DocDoc
    {
        public IVisio.Document doc;
        private Size _pagesize;
        private Rectangle _pagerect;
        private Rectangle _titlerect;
        private Rectangle _bodyrect;
        private int _fontid;
        private TextBlockFormatCells _textblockformat;
        private ParagraphFormatCells _titleParaFmt;
        private ShapeFormatCells _titleFormat;
        private CharacterFormatCells _titleCharFmt;
        private ParagraphFormatCells _bodyParaFmt;
        private CharacterFormatCells _bodyCharFmt;
        private ShapeFormatCells _bodyFormat;

        public DocDoc(IVisio.Application app)
        {
            _pagesize = new VA.Drawing.Size(8.5, 11);
            _pagerect = new VA.Drawing.Rectangle(new VA.Drawing.Point(0, 0), _pagesize);
            _titlerect = new VA.Drawing.Rectangle(_pagerect.UpperLeft.Add(0.5, -1.0), _pagerect.UpperRight.Subtract(0.5, 0.5));
            _bodyrect = new VA.Drawing.Rectangle(_pagerect.LowerLeft.Add(0.5, 0.5), _pagerect.UpperRight.Subtract(0.5, 1.0));

            var docs = app.Documents;
            doc = docs.Add("");

            doc.Subject = "VisioAutomation.Scripting Documenation";
            doc.Title = "VisioAutomation.Scripting Documenation";
            doc.Creator = "";
            doc.Company = "";
            
            var font = doc.Fonts["Segoe UI"];
            _fontid = font.ID;

            _textblockformat = new VA.Text.TextBlockFormatCells();
            _textblockformat.VerticalAlign = 0;

            _titleParaFmt = new VA.Text.ParagraphFormatCells();
            _titleParaFmt.HorizontalAlign = 0;

            _titleFormat = new VA.Format.ShapeFormatCells();
            _titleFormat.LineWeight = 0;
            _titleFormat.LinePattern = 0;

            _titleCharFmt = new VA.Text.CharacterFormatCells();
            _titleCharFmt.Font = _fontid;
            _titleCharFmt.Size = VA.Convert.PointsToInches(15.0);

            _bodyParaFmt = new VA.Text.ParagraphFormatCells();
            _bodyParaFmt.HorizontalAlign = 0;
            _bodyParaFmt.SpacingAfter = VA.Convert.PointsToInches(6.0);

            _bodyCharFmt = new VA.Text.CharacterFormatCells();
            _bodyCharFmt.Font = _fontid;
            _bodyCharFmt.Size = VA.Convert.PointsToInches(8.0);

            _bodyFormat = new VA.Format.ShapeFormatCells();
            _bodyFormat.LineWeight = 0;
            _bodyFormat.LinePattern = 0;
        }

        public void Draw(DocPage xpage)
        {
            var page = doc.Pages.Add();
            page.NameU = xpage.Name;
            VA.PageHelper.SetSize(page, this._pagesize);

            // Draw the shapes
            var titleshape = page.DrawRectangle(_titlerect);
            titleshape.Text = xpage.Title;

            var bodyshape = page.DrawRectangle(_bodyrect);
            bodyshape.Text = xpage.Body;

            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            // Set the ShapeSheet props
            short bodyshape_id = bodyshape.ID16;
            short titleshape_id = titleshape.ID16;
            _textblockformat.Apply(update, titleshape_id);
            this._titleParaFmt.Apply(update, titleshape_id, 0);
            this._titleCharFmt.Apply(update, titleshape_id, 0);
            this._titleFormat.Apply(update, titleshape_id);

            _textblockformat.Apply(update, bodyshape_id);
            this._bodyCharFmt.Apply(update, bodyshape_id, 0);
            this._bodyParaFmt.Apply(update, bodyshape_id, 0);
            this._bodyFormat.Apply(update, bodyshape_id);
            update.Execute(page);
            
        }

        public void Finish()
        {
            // Delete the empty first page
            var first_page = doc.Pages[1];
            first_page.Delete(1);
            first_page = null;

            // set the new first page
            first_page = doc.Pages[1];
            first_page.Activate();

        }

    }

}
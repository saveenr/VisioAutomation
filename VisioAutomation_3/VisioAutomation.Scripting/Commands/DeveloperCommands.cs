using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class DeveloperCommands : SessionCommands
    {
        public DeveloperCommands(Session session) :
            base(session)
        {

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

        public virtual IVisio.Document DrawDocumentationX()
        {
            var app = this.Session.VisioApplication;
            var docs = app.Documents;
            var doc = docs.Add("");

            doc.Subject = "VisioAutomation.Scripting Documenation";
            doc.Title = "VisioAutomation.Scripting Documenation";
            doc.Creator = "";
            doc.Company = "";

            var lines = new List<string>();

            var xxx = typeof (VA.Scripting.Session).GetProperties()
                .Where(
                p => typeof (VA.Scripting.SessionCommands).IsAssignableFrom(p.PropertyType))
                .Select(p=>p.PropertyType)
                .ToList();

            foreach (var mytype in xxx)
            {
                lines.Clear();


                var page = doc.Pages.Add();
                var pagesize = new VA.Drawing.Size(8.5, 11);
                var pagerect = new VA.Drawing.Rectangle(new VA.Drawing.Point(0, 0), pagesize);
                VA.PageHelper.SetSize(page, pagesize);

                // retrieve all public nonstatic methods
                var methods = mytype.GetMethods().Where(m => m.IsPublic).Where(m => !m.IsStatic);

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


                lines.Sort();

                var helpstr = new System.Text.StringBuilder(lines.Select(s => s.Length + 2).Sum());
                foreach (var line in lines)
                {
                    helpstr.Append(line);
                    helpstr.Append("\r\n");
                }

                var titlerect = new VA.Drawing.Rectangle(pagerect.UpperLeft.Add(0.5, -1.0),
                         pagerect.UpperRight.Subtract(0.5, 0.5));

                var titleshape = page.DrawRectangle(titlerect);
                titleshape.Text = mytype.FullName;
                short titleshape_id = titleshape.ID16;

                var bodyrect = new VA.Drawing.Rectangle(pagerect.LowerLeft.Add(0.5, 0.5),
                                                     pagerect.UpperRight.Subtract(0.5, 1.0));

                var bodyshape = page.DrawRectangle(bodyrect);
                bodyshape.Text = helpstr.ToString();
                short bodyshapeid = bodyshape.ID16;

                var textblockformat = new VA.Text.TextBlockFormatCells();
                textblockformat.VerticalAlign = 0;

                var pf = new VA.Text.ParagraphFormatCells();
                pf.HorizontalAlign = 0;

                var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

                textblockformat.Apply(update, titleshape_id);
                textblockformat.Apply(update, bodyshapeid);

                pf.Apply(update, titleshape_id, 0);
                pf.Apply(update, bodyshapeid, 0);

                update.Execute(page);
            }


            return doc;
        }
    }
}
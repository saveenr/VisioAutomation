using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using System.Linq;
using VisioAutomation.Extensions;

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

        public virtual IVisio.Document DrawDocumentation()
        {
            var app = this.Session.VisioApplication;
            var docs = app.Documents;
            var doc = docs.Add("");
            var page = doc.Pages[1];
            var pagesize = new VA.Drawing.Size(8.5, 11);
            var pagerect = new VA.Drawing.Rectangle( new VA.Drawing.Point(0,0),pagesize);
            VA.PageHelper.SetSize(page, pagesize );

            doc.Subject = "VisioAutomation.Scripting Documenation";
            doc.Title= "VisioAutomation.Scripting Documenation";
            doc.Creator = "";
            doc.Company = "";

            var lines = new List<string>();

            var mytype = this.GetType();

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

            pf.Apply(update, titleshape_id, 0 );
            pf.Apply(update, bodyshapeid, 0);

            update.Execute(page);

            return doc;
        }

    }
}
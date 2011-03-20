using System.Collections.Generic;
using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Linq;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        private IVisio.Application m_app;
        private VAS.Session m_scripting_session;
        public bool Debug = false;
        public bool Verbose = false;

        public VisioIPySession()
        {
        }

        public VisioIPySession(IVisio.Application app)
        {
            this.Application = app;
        }

        public VAS.Session ScriptingSession
        {
            get
            {
                if (this.m_scripting_session == null)
                {
                    var scriptingsession = new VAS.Session(this.m_app);
                    scriptingsession.Options = new VisioIPy.VisioIPySessionOptions(this);
                    this.m_scripting_session = scriptingsession;
                }
                else
                {
                    this.m_scripting_session.Application = this.m_app;
                }

                return this.m_scripting_session;
            }
        }

        private void print_app_window_text()
        {
            string visio_window_title = this.ScriptingSession.GetApplicationWindowText();
            this.WriteLine("Window title: \"{0}\"", visio_window_title);
        }

        public void Help()
        {
            var methods = typeof (VisioIPySession).GetMethods()
                .Where(m => m.IsPublic)
                .OrderBy(m=>m.Name);

            foreach (var method in methods)
            {
                var args = method.GetParameters();
                string arg_desc = string.Join(", ", args.Select(a => a.Name));
                System.Console.WriteLine("{0}({1})", method.Name, arg_desc);
            }
        }
    }
}
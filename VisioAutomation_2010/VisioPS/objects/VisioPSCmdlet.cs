using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS
{
    public class VisioPSCmdlet : SMA.Cmdlet
    {
        private static VA.Scripting.Session cached_session;
        internal static ModuleGlobals Globals = new ModuleGlobals();

        public VA.Scripting.Session ScriptingSession
        {
            get
            {
                // if a scripting session is not available create one and cache it
                if (cached_session==null)
                {
                    cached_session = new VA.Scripting.Session(Globals.Application);
                }

                // Must always setup the session output
                // if we try to do this only once per new session then we'll
                // get the "The WriteObject and WriteError methods cannot be
                //     called from outside the overrides of the BeginProcessing
                //     ProcessRecord, and EndProcessing methods, and only
                //     from that same thread."
                // message.

                cached_session.Context = new VisioPSSessionContext(this);
                return cached_session;
            }
        }

        public void WriteVerboseEx(string fmt, params object[] items)
        {
            string s = string.Format(fmt, items);
            this.WriteVerbose(s);
        }
    }
}
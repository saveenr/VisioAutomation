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
                // get this message:
                //
                //    "The WriteObject and WriteError methods cannot be
                //     called from outside the overrides of the BeginProcessing
                //     ProcessRecord, and EndProcessing methods, and only
                //     from that same thread."

                cached_session.Context = new VisioPSSessionContext(this);
                return cached_session;
            }
        }

        public void WriteVerboseEx(string fmt, params object[] items)
        {
            string s = string.Format(fmt, items);
            this.WriteVerbose(s);
        }
        
        protected bool CheckFileExists(string file)
        {
            if (!System.IO.File.Exists(file))
            {
                this.WriteVerboseEx("Filename: {0}",file);
                this.WriteVerboseEx("Abs Filename: {0}", System.IO.Path.GetFullPath(file));
                var exc = new System.IO.FileNotFoundException(file);
                var er = new SMA.ErrorRecord(exc, "FILE_NOT_FOUND", SMA.ErrorCategory.ResourceUnavailable, null);
                this.WriteError(er);
                return false;
            }
            return true;
        }
    }
}
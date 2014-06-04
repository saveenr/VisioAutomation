using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS
{
    public class VisioCmdlet : SMA.Cmdlet
    {
        // this static scripting_session variable is what allows
        // the various visiops cmdlets to share state (for example
        // to share which instance of Visio they are attached to)
        // 
        // To prevent confustion this should be the only static 
        // variable defined in VisioPS
        private static VA.Scripting.Session scripting_session;

        // Attached Visio Application represents the Visio instance
        //
        // that will be used for the cmdlet
        // NOTE that there are three cases - all are valid - to think about:
        // AttachedApplication = null
        // AttachedApplication != null && it is a usable instance
        // AttachedApplication != null && it is an unusable instance. For example
        //                     it might have been manually deleted

        public VA.Scripting.Session ScriptingSession
        {
            get
            {
                // if a scripting session is not available create one and cache it
                // for the lifetime of this cmdlet

                if (scripting_session==null)
                {
                    scripting_session = new VA.Scripting.Session(null);
                }

                // Must always setup the session output
                // if we try to do this only once per new session then we'll
                // get this message:
                //
                //    "The WriteObject and WriteError methods cannot be
                //     called from outside the overrides of the BeginProcessing
                //     ProcessRecord, and EndProcessing methods, and only
                //     from that same thread."

                scripting_session.Context = new VisioPSSessionContext(this);
                return scripting_session;
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
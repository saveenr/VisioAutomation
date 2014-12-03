using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPowerShell
{
    public class VisioCmdlet : SMA.Cmdlet
    {
        // this static _client variable is what allows
        // the various visiops cmdlets to share state (for example
        // to share which instance of Visio they are attached to)
        // 
        // To prevent confustion this should be the only static 
        // variable defined in VisioPS
        private static VA.Scripting.Client _client;

        // Attached Visio Application represents the Visio instance
        //
        // that will be used for the cmdlet
        // NOTE that there are three cases - all are valid - to think about:
        // AttachedApplication = null
        // AttachedApplication != null && it is a usable instance
        // AttachedApplication != null && it is an unusable instance. For example
        //                     it might have been manually deleted

        public VA.Scripting.Client client
        {
            get
            {
                // if a scripting client is not available create one and cache it
                // for the lifetime of this cmdlet

                _client = _client ?? new VA.Scripting.Client(null);
                _client.Context = new VisioPsContext(this);
                return _client;

                // Must always setup the client output
                // if we try to do this only once per new client then we'll
                // get this message:
                //
                //    "The WriteObject and WriteError methods cannot be
                //     called from outside the overrides of the BeginProcessing
                //     ProcessRecord, and EndProcessing methods, and only
                //     from that same thread."

            }
        }

        public void WriteVerbose(string fmt, params object[] items)
        {
            string s = string.Format(fmt, items);
            base.WriteVerbose(s);
        }
        
        protected bool CheckFileExists(string file)
        {
            if (!System.IO.File.Exists(file))
            {
                this.WriteVerbose("Filename: {0}",file);
                this.WriteVerbose("Abs Filename: {0}", System.IO.Path.GetFullPath(file));
                var exc = new System.IO.FileNotFoundException(file);
                var er = new SMA.ErrorRecord(exc, "FILE_NOT_FOUND", SMA.ErrorCategory.ResourceUnavailable, null);
                this.WriteError(er);
                return false;
            }
            return true;
        }
    }
}
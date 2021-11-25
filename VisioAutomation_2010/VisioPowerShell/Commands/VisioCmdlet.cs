

namespace VisioPowerShell.Commands;

public class VisioCmdlet : SMA.Cmdlet
{
    // this static _client variable is what allows
    // the various visiops cmdlets to share state (for example
    // to share which instance of Visio they are attached to)
    // 
    // To prevent confusion this should be the only static 
    // variable defined in VisioPS
    private static VisioScripting.Client _client;

    // Attached Visio Application represents the Visio instance
    //
    // that will be used for the cmdlet
    // NOTE that there are three cases - all are valid - to think about:
    // AttachedApplication = null
    // AttachedApplication != null && it is a usable instance
    // AttachedApplication != null && it is an unusable instance. For example
    //                     it might have been manually deleted

    public VisioScripting.Client Client
    {
        get
        {
            // if a scripting client is not available create one and cache it
            // for the lifetime of this cmdlet

            var ctx = new VisioPsClientContext(this);
            VisioCmdlet._client = VisioCmdlet._client ?? new VisioScripting.Client(null,ctx);
            return VisioCmdlet._client;

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

    protected void _new_app_if_needed()
    {
        if (!this.Client.Application.HasApplication)
        {
            this.Client.Application.NewApplication();
        }
        else
        {
            if (!this.Client.Application.ValidateApplication())
            {
                this.Client.Application.NewApplication();
            }
        }
    }


}
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Close, "VisioDocument")]
    public class Close_VisioDocument: VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter Force;
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter AllDocuments;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            if (this.AllDocuments)
            {
                if (this.Force == false)
                {
                    this.WriteVerbose("Closing All documents requires using the -AllDocuments flag");
                }
                else
                {
                    scriptingsession.Document.CloseAllWithoutSaving();
                }
            }
            else
            {
                scriptingsession.Document.Close(this.Force);
                
            }
        }
    }
}
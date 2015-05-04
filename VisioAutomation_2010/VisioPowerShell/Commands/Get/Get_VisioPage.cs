using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Get, "VisioPage")]
    public class Get_VisioPage : VisioCmdlet
    {
        [SMA.ParameterAttribute(Position=0, Mandatory = false)]
        public string Name=null;

        [SMA.ParameterAttribute(Mandatory = false)] public SMA.SwitchParameter ActivePage;

        protected override void ProcessRecord()
        {
            var application = this.client.VisioApplication;

            if (this.ActivePage)
            {
                var page = this.client.Page.Get();
                this.WriteObject(page);
                return;
            }

            var pages = this.client.Page.GetPagesByName(this.Name);
            this.WriteObject(pages, true);
        }
    }
}
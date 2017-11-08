using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, Nouns.VisioPageLayout)]
    public class SetVisioPageLayout : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)] 
        public VisioScripting.Models.PageOrientation? Orientation = null;
        
        [SMA.Parameter(Mandatory = false)] 
        public string BackgroundPage = null;

        protected override void ProcessRecord()
        {
            if (this.Orientation.HasValue)
            {
                var cmdtarget = this.Client.GetCommandTargetPage();
                var tp = new VisioScripting.Models.TargetPages(cmdtarget.ActivePage);
                this.Client.Page.SetPageOrientation(tp,this.Orientation.Value);
            }

            if (this.BackgroundPage != null)
            {
                this.Client.Page.SetActivePageBackground(this.BackgroundPage);
            }
        }
    }
}
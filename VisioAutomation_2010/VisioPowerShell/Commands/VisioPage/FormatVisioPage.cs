using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Format, Nouns.VisioPage)]
    public class FormatVisioPage: VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)] 
        public VisioScripting.Models.PageOrientation? Orientation = null;
        
        [SMA.Parameter(Mandatory = false)] 
        public string BackgroundPage = null;

        [SMA.Parameter(Mandatory = false)]
        public VisioAutomation.PageLayouts.LayoutBase Layout = null;

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

            if (this.Layout!=null)
            {
                var cmdtarget = this.Client.GetCommandTargetPage();
                var tp = new VisioScripting.Models.TargetPage(cmdtarget.ActivePage);
                this.Client.Page.LayoutPage(tp, this.Layout);
            }
        }
    }
}
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioPageLayout")]
    public class Set_VisioPageLayout : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)] 
        public PageOrientation Orientation = PageOrientation.None;
        
        [SMA.Parameter(Mandatory = false)] 
        public string BackgroundPage = null;

        protected override void ProcessRecord()
        {
            if (this.Orientation != PageOrientation.None)
            {
                this.client.Page.SetOrientation((VA.Pages.PrintPageOrientation)Orientation);
            }

            if (this.BackgroundPage != null)
            {
                this.client.Page.SetBackgroundPage(this.BackgroundPage);
            }
        }
    }
}
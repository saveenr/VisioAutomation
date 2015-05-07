using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Set, "VisioPageLayout")]
    public class Set_VisioPageLayout : VisioCmdlet
    {
        [SMA.ParameterAttribute(Mandatory = false)] 
        public PageOrientation Orientation = PageOrientation.None;
        
        [SMA.ParameterAttribute(Mandatory = false)] 
        public string BackgroundPage = null;

        protected override void ProcessRecord()
        {
            if (this.Orientation != PageOrientation.None)
            {
                this.client.Page.SetOrientation((VisioAutomation.Pages.PrintPageOrientation) this.Orientation);
            }

            if (this.BackgroundPage != null)
            {
                this.client.Page.SetBackgroundPage(this.BackgroundPage);
            }
        }
    }
}
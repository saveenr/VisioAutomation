using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Set, VisioPowerShell.Commands.Nouns.VisioPageLayout)]
    public class Set_VisioPageLayout : VisioCmdlet
    {
        [Parameter(Mandatory = false)] 
        public VisioScripting.Models.PageOrientation? Orientation = null;
        
        [Parameter(Mandatory = false)] 
        public string BackgroundPage = null;

        protected override void ProcessRecord()
        {
            if (this.Orientation.HasValue)
            {
                this.Client.Page.SetOrientation(this.Orientation.Value);
            }

            if (this.BackgroundPage != null)
            {
                this.Client.Page.SetBackgroundPage(this.BackgroundPage);
            }
        }
    }
}
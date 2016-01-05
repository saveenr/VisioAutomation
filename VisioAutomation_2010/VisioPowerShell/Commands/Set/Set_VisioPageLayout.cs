using System.Management.Automation;

namespace VisioPowerShell.Commands.Set
{
    [Cmdlet(VerbsCommon.Set, VisioPowerShell.Nouns.VisioPageLayout)]
    public class Set_VisioPageLayout : VisioCmdlet
    {
        [Parameter(Mandatory = false)] 
        public Model.PageOrientation Orientation = Model.PageOrientation.None;
        
        [Parameter(Mandatory = false)] 
        public string BackgroundPage = null;

        protected override void ProcessRecord()
        {
            if (this.Orientation != Model.PageOrientation.None)
            {
                this.Client.Page.SetOrientation((VisioAutomation.Pages.PrintPageOrientation) this.Orientation);
            }

            if (this.BackgroundPage != null)
            {
                this.Client.Page.SetBackgroundPage(this.BackgroundPage);
            }
        }
    }
}
using System.Management.Automation;
using PageOrientation = VisioScripting.Models.PageOrientation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Set, VisioPowerShell.Nouns.VisioPageLayout)]
    public class Set_VisioPageLayout : VisioCmdlet
    {
        [Parameter(Mandatory = false)] 
        public Models.PageOrientation Orientation = Models.PageOrientation.None;
        
        [Parameter(Mandatory = false)] 
        public string BackgroundPage = null;

        protected override void ProcessRecord()
        {
            if (this.Orientation != Models.PageOrientation.None)
            {
                this.Client.Page.SetOrientation((PageOrientation) this.Orientation);
            }

            if (this.BackgroundPage != null)
            {
                this.Client.Page.SetBackgroundPage(this.BackgroundPage);
            }
        }
    }
}
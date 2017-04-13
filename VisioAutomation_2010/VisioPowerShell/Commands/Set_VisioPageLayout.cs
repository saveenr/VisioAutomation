using System.Management.Automation;
using VisioAutomation.Scripting.Models;
using VisioPowerShell.Models;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Set, VisioPowerShell.Nouns.VisioPageLayout)]
    public class Set_VisioPageLayout : VisioCmdlet
    {
        [Parameter(Mandatory = false)] 
        public PageOrientation Orientation = PageOrientation.None;
        
        [Parameter(Mandatory = false)] 
        public string BackgroundPage = null;

        protected override void ProcessRecord()
        {
            if (this.Orientation != PageOrientation.None)
            {
                this.Client.Page.SetOrientation((PrintPageOrientation) this.Orientation);
            }

            if (this.BackgroundPage != null)
            {
                this.Client.Page.SetBackgroundPage(this.BackgroundPage);
            }
        }
    }
}
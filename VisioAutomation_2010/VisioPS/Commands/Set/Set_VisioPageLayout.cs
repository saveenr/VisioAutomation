using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioPageLayout")]
    public class Set_VisioPageLayout : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)] 
        public PageOrientation Orientation = PageOrientation.None;
        
        [SMA.Parameter(Mandatory = false)] 
        public string BackgroundPage = null;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            if (this.Orientation != PageOrientation.None)
            {
                scriptingsession.Page.SetOrientation((VA.Pages.PrintPageOrientation)Orientation);
            }

            if (this.BackgroundPage != null)
            {
                scriptingsession.Page.SetBackgroundPage(this.BackgroundPage);
            }
        }
    }
}
using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "Page")]
    public class New_Page : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)] public double Width = -1.0;

        [SMA.Parameter(Mandatory = false)] public double Height = -1.0;

        [SMA.Parameter(Mandatory = false)]
        public string Name { get; set; }

        protected override void ProcessRecord()
        {
            var scripting_session = this.ScriptingSession;
            var page = scripting_session.Page.New(null, false);
            Set_PageLayout.set_page_size(scripting_session, Width, Height);
            
            if (this.Name != null)
            {
                scripting_session.Page.SetName(this.Name);
            }
        }
    }
}
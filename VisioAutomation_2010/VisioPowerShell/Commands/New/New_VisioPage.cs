using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioPage")]
    public class New_VisioPage : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)] 
        public double Width = -1.0;
        
        [SMA.Parameter(Mandatory = false)] 
        public double Height = -1.0;

        [SMA.Parameter(Mandatory = false)]
        public string Name { get; set; }

        protected override void ProcessRecord()
        {
            var scripting_session = this.ScriptingSession;
            var page = scripting_session.Page.New(null, false);
            set_page_size(scripting_session, Width, Height);
            
            if (this.Name != null)
            {
                scripting_session.Page.SetName(this.Name);
            }

            this.WriteObject(page);
        }

        public static void set_page_size(VA.Scripting.Session scriptingsession, double width, double height)
        {
            double? w = null;
            double? h = null;

            if (width > 0)
            {
                w = width;
            }

            if (height > 0)
            {
                h = height;
            }

            if (w.HasValue || h.HasValue)
            {
                scriptingsession.Page.SetSize(w, h);
            }
        }
    }
}
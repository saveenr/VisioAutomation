using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioPageLayout")]
    public class Set_VisioPageLayout : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)] public double Width = 0;
        [SMA.Parameter(Mandatory = false)] public double Height = 0;
        [SMA.Parameter(Mandatory = false)] public PageOrientation Orientation = PageOrientation.None;
        [SMA.Parameter(Mandatory = false)] public string BackgroundPage = null;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            set_page_size(scriptingsession, Width, Height);

            if (this.Orientation != PageOrientation.None)
            {
                scriptingsession.Page.SetOrientation((VA.Pages.PrintPageOrientation)Orientation);
            }

            if (this.BackgroundPage != null)
            {
                scriptingsession.Page.SetBackgroundPage(this.BackgroundPage);
            }
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
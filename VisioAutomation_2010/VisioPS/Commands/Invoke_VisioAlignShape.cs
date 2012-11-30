using VA=VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsLifecycle.Invoke, "VisioAlignShape")]
    public class Invoke_VisioAlignShape : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)] public VerticalAlignment Vertical =
            VerticalAlignment.None;

        [SMA.Parameter(Mandatory = false)] public HorizontalAlignment Horizontal
            = HorizontalAlignment.None;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            if (this.Vertical != VerticalAlignment.None)
            {
                scriptingsession.Layout.Align((VA.Drawing.AlignmentVertical)Vertical);
            }
            if (this.Horizontal != HorizontalAlignment.None)
            {
                scriptingsession.Layout.Align((VA.Drawing.AlignmentHorizontal)Horizontal);
            }
        }
    }
}
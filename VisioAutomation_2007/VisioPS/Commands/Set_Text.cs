using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Set", "Text")]
    public class Set_Text : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string[] Text { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.Text.SetText(Text);
        }
    }
}

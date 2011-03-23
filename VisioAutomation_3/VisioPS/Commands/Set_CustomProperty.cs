using VAS = VisioAutomation.Scripting;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Set", "CustomProperty")]
    public class Set_CustomProperty : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string Name { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public string Value { get; set; }

        [SMA.Parameter(Mandatory = false)] public string Label;

        [SMA.Parameter(Mandatory = false)] public string Prompt;

        [SMA.Parameter(Mandatory = false)] public int LangID = -1;

        [SMA.Parameter(Mandatory = false)] public int SortKey = -1;

        [SMA.Parameter(Mandatory = false)] public
            VA.CustomProperties.Format Type = VA.CustomProperties.Format.String;

        [SMA.Parameter(Mandatory = false)] public int Verify = -1;

        protected override void ProcessRecord()
        {
            var cp = new VA.CustomProperties.CustomPropertyCells();
            cp.Value = this.Value;
            if (this.Label != null)
            {
                cp.Label = this.Label;
            }

            if (this.LangID >= 0)
            {
                cp.LangId = this.LangID;
            }

            if (this.Prompt != null)
            {
                cp.Prompt = this.Prompt;
            }

            if (this.SortKey >= 0)
            {
                cp.SortKey = this.SortKey;
            }

            cp.Type = (int) this.Type;

            if (this.Verify >= 0)
            {
                cp.Verify = this.Verify;
            }

            var scriptingsession = this.ScriptingSession;
            scriptingsession.CustomProp.SetCustomProperty(this.Name , cp);
        }
    }
}
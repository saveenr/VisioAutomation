using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Set", "UserDefinedCell")]
    public class Set_UserDefinedCell : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string Name { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public string Value { get; set; }

        [SMA.Parameter(Mandatory = false)] public string Prompt;

        protected override void ProcessRecord()
        {
            var userprop = new VA.UserDefinedCells.UserDefinedCell(this.Name, this.Value);
            if (this.Prompt != null)
            {
                userprop.Prompt = this.Prompt;
            }

            var scriptingsession = this.ScriptingSession;
            scriptingsession.UserDefinedCell.SetUserDefinedCell(userprop);
        }
    }
}
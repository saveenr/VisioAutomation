using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Move", "VisioTextBox")]
    public class Move_VisioTextBox : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)] public TextBoxPosition Position;

        protected override void ProcessRecord()
        {
            if (this.Position == TextBoxPosition.BottomOfShape)
            {
                var scriptingsession = this.ScriptingSession;
                scriptingsession.Text.MoveTextToBottom();
            }
        }
    }
}
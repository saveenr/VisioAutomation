using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Move", "TextBox")]
    public class Move_TextBox : VisioPSCmdlet
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
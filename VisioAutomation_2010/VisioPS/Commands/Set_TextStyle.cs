using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{    
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "TextStyle")]
    public class Set_TextStyle : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = true, Position=0)] public string Name;
        [SMA.Parameter(Mandatory = false)] public string FontName;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            scriptingsession.Text.SetStyleProperties(this.Name, this.FontName);
        }
    }
}
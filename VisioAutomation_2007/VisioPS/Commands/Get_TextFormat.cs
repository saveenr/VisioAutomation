using VAS =VisioAutomation.Scripting;
using SMA = System.Management.Automation;
using VA=VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Get", "TextFormat")]
    public class Get_TextFormat : VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var t = scriptingsession.Text.GetTextFormat();
            this.WriteObject(t);
        }
    }

    [SMA.Cmdlet("Set", "TextFormat")]
    public class Set_TextFormat : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public VA.Text.CharacterFormatCells CharacterFormatCells { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public VA.Text.ParagraphFormatCells ParagraphFormatCells { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.Text.SetTextFormat(this.CharacterFormatCells,this.ParagraphFormatCells);
        }
    }

    [SMA.Cmdlet("New", "CharacterFormatCells")]
    public class New_CharacterFormatCells : VisioPSCmdlet
    {

        protected override void ProcessRecord()
        {
            var cells = new VA.Text.CharacterFormatCells();
            this.WriteObject(cells);
        }
    }

    [SMA.Cmdlet("New", "ParagraphFormatCells")]
    public class ParagraphFormatCells : VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var cells = new VA.Text.ParagraphFormatCells();
            this.WriteObject(cells);
        }
    }
}
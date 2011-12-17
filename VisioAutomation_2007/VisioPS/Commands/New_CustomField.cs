using VAS=VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("New", "CustomField")]
    public class New_CustomField: VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public int Start;

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public int End;
        
        [SMA.Parameter(Position = 2, Mandatory = true)]
        public string Formula;

        [SMA.Parameter(Position = 3, Mandatory = false)]
        IVisio.VisFieldFormats Format = IVisio.VisFieldFormats.visFmtNumGenNoUnits;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.Text.InsertCustomField(Start, End, Formula, Format);
        }
    }
}
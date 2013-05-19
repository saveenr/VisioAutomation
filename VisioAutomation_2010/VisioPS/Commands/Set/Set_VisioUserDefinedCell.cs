using System.Collections.Generic;
using VA = VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioUserDefinedCell")]
    public class Set_VisioUserDefinedCell : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string Name { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public string Value { get; set; }

        [SMA.Parameter(Mandatory = false)] public string Prompt;

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes; 

        protected override void ProcessRecord()
        {
            var userprop = new VA.UserDefinedCells.UserDefinedCell(this.Name, this.Value);
            if (this.Prompt != null)
            {
                userprop.Prompt = this.Prompt;
            }

            var scriptingsession = this.ScriptingSession;
            scriptingsession.UserDefinedCell.Set(this.Shapes,userprop);
        }
    }
}
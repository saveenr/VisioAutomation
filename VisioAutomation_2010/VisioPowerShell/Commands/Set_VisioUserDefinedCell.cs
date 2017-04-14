using System.Management.Automation;
using VisioAutomation.Scripting.Models;
using VisioAutomation.Shapes;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Set, VisioPowerShell.Nouns.VisioUserDefinedCell)]
    public class Set_VisioUserDefinedCell : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        public string Name { get; set; }

        [Parameter(Position = 1, Mandatory = true)]
        public string Value { get; set; }

        [Parameter(Mandatory = false)] 
        public string Prompt;

        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes; 

        protected override void ProcessRecord()
        {
            var targets = new TargetShapes(this.Shapes);
            var userprop = new UserDefinedCellCells(this.Name, this.Value);
            if (this.Prompt != null)
            {
                userprop.Prompt = this.Prompt;
            }

            this.Client.UserDefinedCell.Set(targets, userprop);
        }
    }
}
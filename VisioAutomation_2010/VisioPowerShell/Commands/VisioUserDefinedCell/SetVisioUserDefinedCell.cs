using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, Nouns.VisioUserDefinedCell)]
    public class SetVisioUserDefinedCell : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string Name { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public string Value { get; set; }

        [SMA.Parameter(Mandatory = false)] 
        public string Prompt;

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes; 

        protected override void ProcessRecord()
        {
            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);
            var userprop = new VisioScripting.Models.UserDefinedCell(this.Name, this.Value);
            if (this.Prompt != null)
            {
                userprop.Cells.Prompt = this.Prompt;
            }

            this.Client.UserDefinedCell.SetUserDefinedCell(targets, userprop);
        }
    }
}
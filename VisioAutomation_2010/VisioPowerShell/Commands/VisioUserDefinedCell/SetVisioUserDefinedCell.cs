

namespace VisioPowerShell.Commands.VisioUserDefinedCell
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

        // CONTEXT:SHAPES 
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shape; 

        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shape);
            var udcell = new VisioAutomation.Shapes.UserDefinedCellCells();

            if (this.Value != null)
            {
                udcell.Value = this.Value;
            }
            if (this.Prompt != null)
            {
                udcell.Prompt = this.Prompt;
            }

            this.Client.UserDefinedCell.SetUserDefinedCell(targetshapes, this.Name, udcell);
        }
    }
}
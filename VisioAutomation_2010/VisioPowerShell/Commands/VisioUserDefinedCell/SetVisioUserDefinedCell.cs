using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

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

            // Encode -Value via SetString and -Prompt via EncodeValue: both
            // fields are Visio formulas, not literal strings. Without this,
            // a typical -Value 'foo' arg would reach UserDefinedCellHelper.Set
            // unencoded, which #144's detect-and-rethrow then surfaces as
            // ArgumentException ("Visio rejected the formula ... use SetString
            // ... see #144"). Unlike Set-VisioCustomProperty, the UDC
            // VisioScripting layer has no EncodeValues backstop, so the
            // encoding has to happen here.
            if (this.Value != null)
            {
                udcell.SetString(this.Value);
            }
            if (this.Prompt != null)
            {
                udcell.Prompt = VisioAutomation.Core.CellValue.EncodeValue(this.Prompt);
            }

            this.Client.UserDefinedCell.SetUserDefinedCell(targetshapes, this.Name, udcell);
        }
    }
}
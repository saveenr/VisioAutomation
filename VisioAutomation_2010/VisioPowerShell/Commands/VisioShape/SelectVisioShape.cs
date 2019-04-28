using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.Select, Nouns.VisioShape)]
    public class SelectVisioShape : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true, Position = 0, ParameterSetName = "SelectByShapes")]
        public IVisio.Shape[] Shapes { get; set; }
       
       
        [SMA.Parameter(Mandatory = true, Position=0, ParameterSetName = "SelectByOperation")] 
        public VisioScripting.Models.SelectionOperation Operation { get; set; }

        protected override void ProcessRecord()
        {
            if (this.Shapes !=null)
            {
                this.Client.Selection.SelectShapes(VisioScripting.TargetWindow.Auto, this.Shapes);
                return;
            }

            if (this.Operation == VisioScripting.Models.SelectionOperation.All)
            {
                this.Client.Selection.SelectAllShapes(VisioScripting.TargetWindow.Auto);
            }
            else if (this.Operation == VisioScripting.Models.SelectionOperation.None)
            {
                this.Client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            }
            else if (this.Operation == VisioScripting.Models.SelectionOperation.Invert)
            {
                this.Client.Selection.InvertSelection(VisioScripting.TargetWindow.Auto);
            }
        }
    }
}
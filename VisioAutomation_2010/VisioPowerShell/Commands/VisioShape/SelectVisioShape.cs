using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.Select, Nouns.VisioShape)]
    public class SelectVisioShape : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true, Position = 0, ParameterSetName = "SelectByShapes")]
        public IVisio.Shape[] Shapes { get; set; }

        [SMA.Parameter(Mandatory = true, Position = 0, ParameterSetName = "SelectByOperation_SelectAll")]
        public SMA.SwitchParameter All;

        [SMA.Parameter(Mandatory = true, Position = 0, ParameterSetName = "SelectByOperation_None")]
        public SMA.SwitchParameter None;

        [SMA.Parameter(Mandatory = true, Position = 0, ParameterSetName = "SelectByOperation_Invert")]
        public SMA.SwitchParameter Invert;

        protected override void ProcessRecord()
        {
            if (this.Shapes !=null)
            {
                this.Client.Selection.SelectShapes(VisioScripting.TargetWindow.Auto, this.Shapes);
                return;
            }

            if (this.All)
            {
                this.Client.Selection.SelectAllShapes(VisioScripting.TargetWindow.Auto);
            }
            else if (this.None)
            {
                this.Client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            }
            else if (this.Invert)
            {
                this.Client.Selection.InvertSelection(VisioScripting.TargetWindow.Auto);
            }
        }
    }
}
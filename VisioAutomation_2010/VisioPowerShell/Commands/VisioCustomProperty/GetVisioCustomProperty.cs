using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;
using VASS = VisioAutomation.ShapeSheet;

namespace VisioPowerShell.Commands.VisioCustomProperty
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioCustomProperty)]
    public class GetVisioCustomProperty : VisioCmdlet
    {
        // CONTEXT:SHAPES
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shape;
        
        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shape);
            var type = VisioAutomation.Core.CellValueType.Formula;
            var dicof_shape_to_cpdic = this.Client.CustomProperty.GetCustomPropertiesAsShapeDictionary(targetshapes, type);
            this.WriteObject(dicof_shape_to_cpdic);                
        }
    }
}
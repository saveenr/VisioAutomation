using VisioPowerShell.Models;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioCustomProperty
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioCustomProperty)]
    public class GetVisioCustomProperty : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;
        
        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shapes);
            var dicof_shape_to_cpdic = this.Client.CustomProperty.GetCustomProperties(targetshapes);
            this.WriteObject(dicof_shape_to_cpdic);                
        }
    }
}
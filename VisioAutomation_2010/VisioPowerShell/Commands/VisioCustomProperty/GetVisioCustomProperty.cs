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

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter GetCells;

        protected override void ProcessRecord()
        {
            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);
            var dicof_shape_to_cpdic = this.Client.CustomProperty.GetCustomProperties(targets);

            if (this.GetCells)
            {
                this.WriteObject(dicof_shape_to_cpdic);                
                return;
            }

            foreach (var shape_cppdic_pair in dicof_shape_to_cpdic)
            {
                var shape = shape_cppdic_pair.Key;
                var cpdic = shape_cppdic_pair.Value;
                int shapeid = shape.ID;
                foreach (var propname_propcells_pair in cpdic)
                {
                    string propname = propname_propcells_pair.Key;
                    var propcells = propname_propcells_pair.Value;
                    var cpo = new CustomProperty(shapeid, propname, propcells);
                    this.WriteObject(cpo);
                }
            }
        }
    }
}
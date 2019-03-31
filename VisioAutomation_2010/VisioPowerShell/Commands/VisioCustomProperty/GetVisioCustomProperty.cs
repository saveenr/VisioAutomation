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
            var targetshapes = new VisioScripting.Models.TargetShapes(this.Shapes);
            var dicof_shape_to_cpdic = this.Client.CustomProperty.GetCustomProperties(targetshapes);

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
                foreach (var cpname_cpcells_pair in cpdic)
                {
                    string cpname = cpname_cpcells_pair.Key;
                    var cpcells = cpname_cpcells_pair.Value;
                    var cp_obj = new CustomProperty(shapeid, cpname, cpcells);
                    this.WriteObject(cp_obj);
                }
            }
        }
    }
}
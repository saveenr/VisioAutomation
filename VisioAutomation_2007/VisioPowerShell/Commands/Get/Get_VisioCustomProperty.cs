using VisioAutomation.Shapes.CustomProperties;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioCustomProperty")]
    public class Get_VisioCustomProperty : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter GetCells;

        protected override void ProcessRecord()
        {
            var dic = this.client.CustomProp.Get(this.Shapes);

            if (this.GetCells)
            {
                this.WriteObject(dic);                
                return;
            }

            foreach (var shape_propdic_pair in dic)
            {
                var shape = shape_propdic_pair.Key;
                var propdic = shape_propdic_pair.Value;
                int shape_id = shape.ID;
                foreach (var propname_propcells_pair in propdic)
                {
                    string propname = propname_propcells_pair.Key;
                    var propcells = propname_propcells_pair.Value;
                    var cpf = new CustomPropertyValues(shape_id, propname, propcells);
                    this.WriteObject(cpf);
                }
            }
        }
    }
}
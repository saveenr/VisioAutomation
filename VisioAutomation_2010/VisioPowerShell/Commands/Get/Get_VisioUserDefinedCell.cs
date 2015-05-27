using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Get
{
    [Cmdlet(VerbsCommon.Get, "VisioUserDefinedCell")]
    public class Get_VisioUserDefinedCell : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        [Parameter(Mandatory = false)]
        public SwitchParameter GetCells;

        protected override void ProcessRecord()
        {
            var dic = this.client.UserDefinedCell.Get(this.Shapes);

            if (this.GetCells)
            {
                this.WriteObject(dic);
                return;
            }

            foreach (var kv in dic)
            {
                int shapeid = kv.Key.ID;
                foreach (var udc in kv.Value)
                {
                    var udcell_vals = new Model.UserDefinedCellvalues(shapeid, udc.Name, udc.Value.Formula.Value,udc.Prompt.Formula.Value);
                    this.WriteObject(udcell_vals);
                }
            }
        }
    }
}
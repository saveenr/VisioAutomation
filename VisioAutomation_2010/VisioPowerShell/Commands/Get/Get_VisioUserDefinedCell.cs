using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Get, "VisioUserDefinedCell")]
    public class Get_VisioUserDefinedCell : VisioCmdlet
    {
        [SMA.ParameterAttribute(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        [SMA.ParameterAttribute(Mandatory = false)]
        public SMA.SwitchParameter GetCells;

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
                    var udcell_vals = new UserDefinedCellvalues(shapeid, udc.Name, udc.Value.Formula.Value,udc.Prompt.Formula.Value);
                    this.WriteObject(udcell_vals);
                }
            }
        }
    }
}
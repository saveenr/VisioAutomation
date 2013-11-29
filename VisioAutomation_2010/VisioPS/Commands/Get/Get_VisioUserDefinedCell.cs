using System.Collections.Generic;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioUserDefinedCell")]
    public class Get_VisioUserDefinedCell : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter GetCells;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var dic = scriptingsession.UserDefinedCell.Get(this.Shapes);

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
                    var rec = new UserDefinedCellRecord();
                    rec.ShapeID = shapeid;
                    rec.Name = udc.Name;
                    rec.Value = udc.Value.Formula.Value;
                    rec.Prompt = udc.Prompt.Formula.Value;
                    this.WriteObject(rec);
                }
            }
        }
    }

    public class UserDefinedCellRecord
    {
        public int ShapeID;
        public string Name { get; set; }
        public string Value { get; set; }
        public string Prompt { get; set; }
    }
}
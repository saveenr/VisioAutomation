using System.Collections.Generic;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioCustomProperty")]
    public class Get_VisioCustomProperty : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter GetCells;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var dic = scriptingsession.CustomProp.Get(this.Shapes);

            if (this.GetCells)
            {
                this.WriteObject(dic);                
            }
            else
            {
                var dt = new System.Data.DataTable("CustomProperties");
                dt.Columns.Add("ShapeId");
                dt.Columns.Add("Name");
                dt.Columns.Add("Value");
                dt.Columns.Add("Format");
                dt.Columns.Add("Invisible");
                dt.Columns.Add("Label");
                dt.Columns.Add("LangId");
                dt.Columns.Add("Prompt");
                dt.Columns.Add("SortKey");
                dt.Columns.Add("Type");
                dt.Columns.Add("Verify");

                foreach (var shape_propdic_pair in dic)
                {
                    var shape = shape_propdic_pair.Key;
                    var propdic = shape_propdic_pair.Value;
                    foreach (var propname_propcells_pair in propdic)
                    {
                        string propname = propname_propcells_pair.Key;
                        var propcells = propname_propcells_pair.Value;

                        dt.Rows.Add(
                            shape.ID, 
                            propname, 
                            propcells.Value.Formula, 
                            propcells.Format.Formula, 
                            propcells.Invisible.Formula,
                            propcells.Label.Formula,
                            propcells.LangId.Formula, 
                            propcells.Prompt.Formula, 
                            propcells.SortKey.Formula,
                            propcells.Type.Formula,
                            propcells.Verify.Formula);
                    }
                }

                this.WriteObject(dt);
            }
        }

    }
}
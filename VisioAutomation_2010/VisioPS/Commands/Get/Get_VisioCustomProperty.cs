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
                return;
            }

            foreach (var shape_propdic_pair in dic)
            {
                var shape = shape_propdic_pair.Key;
                var propdic = shape_propdic_pair.Value;
                foreach (var propname_propcells_pair in propdic)
                {
                    string propname = propname_propcells_pair.Key;
                    var propcells = propname_propcells_pair.Value;

                    var cpf = new CustomPropertyFormulas();
                    cpf.ShapeID = shape.ID;
                    cpf.Name = propname;
                    cpf.Value = propcells.Value.Formula.Value;
                    cpf.Value = propcells.Format.Formula.Value;
                    cpf.Value = propcells.Invisible.Formula.Value;
                    cpf.Value = propcells.Label.Formula.Value;
                    cpf.Value = propcells.LangId.Formula.Value;
                    cpf.Value =  propcells.Prompt.Formula.Value; 
                    cpf.Value =  propcells.SortKey.Formula.Value;
                    cpf.Value = propcells.Type.Formula.Value;
                    cpf.Value = propcells.Verify.Formula.Value;

                    this.WriteObject(cpf);
                }
            }
        }
    }

    public class CustomPropertyFormulas
    {
        public int ShapeID;
        public string Name;
        public string Value;
        public string Format;
        public string Invisible;
        public string Label;
        public string LangId;
        public string Prompt;
        public string SortKey;
        public string Type;
        public string Verify;

        public CustomPropertyFormulas()
        {           
        }
    }
}
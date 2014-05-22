using VA = VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioCustomProperty")]
    public class Set_VisioCustomProperty : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "HashTable")]
        public System.Collections.Hashtable HashTable{ get; set; }
        
        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "NonHashTable")]
        public string Name { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true, ParameterSetName = "NonHashTable")]
        public string Value { get; set; }

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NonHashTable")]
        public string Label { get; set; }

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NonHashTable")]
        public string Prompt { get; set; }

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NonHashTable")]
        public int LangId { get; set; }

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NonHashTable")]
        public int SortKey { get; set; }

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NonHashTable")]
        public int Type { get; set; }

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NonHashTable")]
        public int Verify { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        public Set_VisioCustomProperty()
        {
            Verify = -1;
            Type = 0;
            SortKey = -1;
            LangId = -1;
        }

        protected override void ProcessRecord()
        {
            if (this.HashTable != null)
            {
                SetFromHashTable();
            }
            else
            {
                SetFromParameters();
            }
        }

        private void SetFromParameters()
        {
            var cp = new VA.Shapes.CustomProperties.CustomPropertyCells();
            cp.Value = this.Value;
            if (this.Label != null)
            {
                cp.Label = this.Label;
            }

            if (this.LangId >= 0)
            {
                cp.LangId = this.LangId;
            }

            if (this.Prompt != null)
            {
                cp.Prompt = this.Prompt;
            }

            if (this.SortKey >= 0)
            {
                cp.SortKey = this.SortKey;
            }

            cp.Type = (int) this.Type;


            var scriptingsession = this.ScriptingSession;
            scriptingsession.CustomProp.Set(this.Shapes, this.Name, cp);
        }

        private void SetFromHashTable()
        {
            if (this.HashTable.Count < 1)
            {
                return;
            }

            foreach (object key in this.HashTable.Keys)
            {
                if (!(key is string))
                {
                    string msg = string.Format("Property Names must be strings");
                    throw new System.ArgumentOutOfRangeException(msg);
                }

                string key_string = (string) key;

                object value = this.HashTable[key];
                var cp = VA.Shapes.CustomProperties.CustomPropertyCells.FromValue(value);
                var scriptingsession = this.ScriptingSession;
                scriptingsession.CustomProp.Set(this.Shapes, key_string, cp);
            }
        }
    }
}
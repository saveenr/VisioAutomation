using VA = VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioCustomProperty")]
    public class Set_VisioCustomProperty : VisioCmdlet
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
        public string Format { get; set; }

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NonHashTable")]
        public string Prompt { get; set; }

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NonHashTable")] 
        public int LangId = -1;

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NonHashTable")] 
        public int SortKey = -1;

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NonHashTable")] 
        public int Type = 0;

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NonHashTable")] 
        public int Ask = -1;

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NonHashTable")] 
        public int Calendar = -1;

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NonHashTable")] 
        public int Invisible = -1;

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

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

            if (this.Format != null)
            {
                cp.Format = this.Format;
            }

            if (this.Prompt!= null)
            {
                cp.Prompt = this.Prompt;
            }

            if (this.LangId >= 0)
            {
                cp.LangId = this.LangId;
            }

            if (this.SortKey >= 0)
            {
                cp.SortKey = this.SortKey;
            }

            if (this.Type >= 0)
            {
                cp.Type = this.Type;
            }

            if (this.Ask>= 0)
            {
                cp.Ask = this.Ask;
            }

            if (this.Calendar >= 0)
            {
                cp.Calendar = this.Calendar;
            }

            if (this.Invisible >= 0)
            {
                cp.Invisible = this.Invisible;
            }

            this.client.CustomProp.Set(this.Shapes, this.Name, cp);
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
                this.client.CustomProp.Set(this.Shapes, key_string, cp);
            }
        }
    }
}
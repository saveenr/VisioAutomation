using System;
using System.Collections;
using System.Management.Automation;
using VACUSTPROP = VisioAutomation.Shapes.CustomProperties;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Set
{
    [Cmdlet(VerbsCommon.Set, "VisioCustomProperty")]
    public class Set_VisioCustomProperty : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true, ParameterSetName = "HashTable")]
        public Hashtable HashTable{ get; set; }
        
        [Parameter(Position = 0, Mandatory = true, ParameterSetName = "NonHashTable")]
        public string Name { get; set; }

        [Parameter(Position = 1, Mandatory = true, ParameterSetName = "NonHashTable")]
        public string Value { get; set; }

        [Parameter(Mandatory = false, ParameterSetName = "NonHashTable")]
        public string Label { get; set; }

        [Parameter(Mandatory = false, ParameterSetName = "NonHashTable")]
        public string Format { get; set; }

        [Parameter(Mandatory = false, ParameterSetName = "NonHashTable")]
        public string Prompt { get; set; }

        [Parameter(Mandatory = false, ParameterSetName = "NonHashTable")] 
        public int LangId = -1;

        [Parameter(Mandatory = false, ParameterSetName = "NonHashTable")] 
        public int SortKey = -1;

        [Parameter(Mandatory = false, ParameterSetName = "NonHashTable")] 
        public int Type = 0;

        [Parameter(Mandatory = false, ParameterSetName = "NonHashTable")] 
        public int Ask = -1;

        [Parameter(Mandatory = false, ParameterSetName = "NonHashTable")] 
        public int Calendar = -1;

        [Parameter(Mandatory = false, ParameterSetName = "NonHashTable")] 
        public int Invisible = -1;

        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            if (this.HashTable != null)
            {
                this.SetFromHashTable();
            }
            else
            {
                this.SetFromParameters();
            }
        }

        private void SetFromParameters()
        {
            var cp = new VACUSTPROP.CustomPropertyCells();

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
                    string msg = "Property Names must be strings";
                    throw new ArgumentOutOfRangeException(msg);
                }

                string key_string = (string) key;

                object value = this.HashTable[key];
                var cp = VACUSTPROP.CustomPropertyCells.FromValue(value);
                this.client.CustomProp.Set(this.Shapes, key_string, cp);
            }
        }
    }
}
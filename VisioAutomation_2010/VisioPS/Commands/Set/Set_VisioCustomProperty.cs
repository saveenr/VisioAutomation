using System;
using System.Collections.Generic;
using VisioAutomation.Shapes.CustomProperties;
using VA = VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioCustomProperty")]
    public class Set_VisioCustomProperty : VisioPS.VisioPSCmdlet
    {
        private int _LangID = -1;
        private int _sortKey = -1;
        private int _type = 0; // 0 = string
        private int _verify = -1;

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
        public int LangId
        {
            get { return _LangID; }
            set { _LangID = value; }
        }

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NonHashTable")]
        public int SortKey
        {
            get { return _sortKey; }
            set { _sortKey = value; }
        }

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NonHashTable")]
        public int Type
        {
            get { return _type; }
            set { _type = value; }
        }

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NonHashTable")]
        public int Verify
        {
            get { return _verify; }
            set { _verify = value; }
        }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            if (this.HashTable != null)
            {
                PerformHashTable();
            }
            else
            {
                PerformNonHashTable();
            }
        }

        private void PerformNonHashTable()
        {
            var cp = new CustomPropertyCells();
            cp.Value = this.Value;
            if (this.Label != null)
            {
                cp.Label = this.Label;
            }

            if (this._LangID >= 0)
            {
                cp.LangId = this._LangID;
            }

            if (this.Prompt != null)
            {
                cp.Prompt = this.Prompt;
            }

            if (this._sortKey >= 0)
            {
                cp.SortKey = this._sortKey;
            }

            cp.Type = (int) this._type;


            var scriptingsession = this.ScriptingSession;
            scriptingsession.CustomProp.Set(this.Shapes, this.Name, cp);
        }

        private void PerformHashTable()
        {
            if (this.HashTable.Count < 1)
            {
                return;
            }

            foreach (object  key in this.HashTable.Keys)
            {
                if (!(key is string))
                {
                    string msg = string.Format("Property Names must be strings");
                    throw new System.ArgumentOutOfRangeException(msg);
                }

                string key_string = (string) key;

                object value = this.HashTable[key];
                var cp = CustomPropertyCells.FromValue(value);
                var scriptingsession = this.ScriptingSession;
                scriptingsession.CustomProp.Set(this.Shapes, key_string, cp);
            }
        }
    }
}
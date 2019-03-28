using System;
using System.Collections;
using SMA = System.Management.Automation;
using VisioAutomation.Shapes;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, Nouns.VisioCustomProperty)]
    public class SetVisioCustomProperty : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "HashTable")]
        public Hashtable Hashtable{ get; set; }
        
        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "NonHashTable")]
        public string Name { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true, ParameterSetName = "NonHashTable")]
        public object Value { get; set; }

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NonHashTable")]
        public int Type = 0;

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
        public int Ask = -1;

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NonHashTable")] 
        public int Calendar = -1;

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NonHashTable")] 
        public int Invisible = -1;

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            if (this.Hashtable != null)
            {
                this._set_from_hash_table();
            }
            else
            {
                this._set_from_parameters();
            }
        }

        private void _set_from_parameters()
        {
            // this will set .Value and automatically set
            // .Type as needed.
            var cp = _create_cust_prop_from_object(this.Value);

            // The user can override .Type if desired
            if (this.Type >= 0)
            {
                cp.Type = this.Type;
            }

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
                cp.LangID = this.LangId;
            }

            if (this.SortKey >= 0)
            {
                cp.SortKey = this.SortKey;
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

            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);
            this.Client.CustomProperty.SetCustomProperty(targets, this.Name, cp);
        }

        private void _set_from_hash_table()
        {
            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);

            if (this.Hashtable.Count < 1)
            {
                return;
            }

            foreach (object key in this.Hashtable.Keys)
            {
                if (!(key is string))
                {
                    string msg = "Property Names must be strings";
                    throw new ArgumentOutOfRangeException(msg);
                }

                string key_string = (string) key;

                object value = this.Hashtable[key];
                var cp = _create_cust_prop_from_object(value);
                this.Client.CustomProperty.SetCustomProperty(targets, key_string, cp);
            }
        }

        private static CustomPropertyCells _create_cust_prop_from_object(object value)
        {
            if (value is string value_str)
            {
                return new CustomPropertyCells(value_str);
            }
            else if (value is int value_int)
            {
                return new CustomPropertyCells(value_int);
            }
            else if (value is double value_double)
            {
                return new CustomPropertyCells(value_double);
            }
            else if (value is float value_float)
            {
                return new CustomPropertyCells(value_float);
            }
            else if (value is bool value_bool)
            {
                return new CustomPropertyCells(value_bool);
            }
            else if (value is System.DateTime value_datetime)
            {
                return new CustomPropertyCells(value_datetime);
            }
            else if (value is VisioAutomation.ShapeSheet.CellValueLiteral value_cvl)
            {
                return new CustomPropertyCells(value_cvl);
            }

            var value_type = value.GetType();
            string msg = string.Format("Unsupported type for value \"{0}\" of type \"{1}\"", value, value_type.Name);
            throw new System.ArgumentException(msg);
        }
    }
}
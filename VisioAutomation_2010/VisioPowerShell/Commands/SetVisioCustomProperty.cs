using System;
using System.Collections;
using System.Management.Automation;
using VisioAutomation.Shapes;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Set, VisioPowerShell.Commands.Nouns.VisioCustomProperty)]
    public class SetVisioCustomProperty : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true, ParameterSetName = "HashTable")]
        public Hashtable Hashtable{ get; set; }
        
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
            if (this.Hashtable != null)
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
            var cp = new CustomPropertyCells();

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
                cp.LangID = this.LangId;
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

            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);
            this.Client.CustomProperty.Set(targets, this.Name, cp);
        }

        private void SetFromHashTable()
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
                var cp = CreateCustPropFromObject(value);
                this.Client.CustomProperty.Set(targets, key_string, cp);
            }
        }

        private static CustomPropertyCells CreateCustPropFromObject(object value)
        {
            if (value is string value_str)
            {
                return CustomPropertyCells.Create(value_str);
            }
            else if (value is int value_int)
            {
                return CustomPropertyCells.Create(value_int);
            }
            else if (value is double value_double)
            {
                return CustomPropertyCells.Create(value_double);
            }
            else if (value is float value_float)
            {
                return CustomPropertyCells.Create(value_float);
            }
            else if (value is bool value_bool)
            {
                return CustomPropertyCells.Create(value_bool);
            }
            else if (value is System.DateTime value_datetime)
            {
                return CustomPropertyCells.Create(value_datetime);
            }
            else if (value is VisioAutomation.ShapeSheet.CellValueLiteral value_cvl)
            {
                return CustomPropertyCells.Create(value_cvl);
            }

            var value_type = value.GetType();
            string msg = string.Format("Unsupported type for value \"{0}\" of type \"{1}\"", value, value_type.Name);
            throw new System.ArgumentException(msg);
        }

    }
}
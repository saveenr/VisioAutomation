using VisioAutomation.Shapes;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioCustomProperty
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, Nouns.VisioCustomProperty)]
    public class SetVisioCustomProperty : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string Name { get; set; }


        [SMA.Parameter(Position = 1, Mandatory = true, ParameterSetName = "Cells")]
        public VisioAutomation.Shapes.CustomPropertyCells Cells { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true, ParameterSetName = "NamedProperties")]
        public object Value { get; set; }

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NamedProperties")]
        public int Type = 0;

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NamedProperties")]
        public string Label { get; set; }

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NamedProperties")]
        public string Format { get; set; }

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NamedProperties")]
        public string Prompt { get; set; }

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NamedProperties")] 
        public int LangId = -1;

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NamedProperties")] 
        public int SortKey = -1;

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NamedProperties")] 
        public int Ask = -1;

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NamedProperties")] 
        public int Calendar = -1;

        [SMA.Parameter(Mandatory = false, ParameterSetName = "NamedProperties")] 
        public int Invisible = -1;

        // CONTEXT:SHAPES
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shape;

        protected override void ProcessRecord()
        {
            if (this.Cells != null)
            {
                this._set_from_cells();
            }
            else
            {
                this._set_from_namedproperties();
            }
        }

        private void _set_from_namedproperties()
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

            var targetshapes = new VisioScripting.TargetShapes(this.Shape);
            this.Client.CustomProperty.SetCustomProperty(targetshapes, this.Name, cp);
        }

        private void _set_from_cells()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shape);
            this.Client.CustomProperty.SetCustomProperty(targetshapes, this.Name, this.Cells);
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
            else if (value is VisioAutomation.Core.CellValue value_cvl)
            {
                return new CustomPropertyCells(value_cvl);
            }

            var value_type = value.GetType();
            string msg = string.Format("Unsupported type for value \"{0}\" of type \"{1}\"", value, value_type.Name);
            throw new System.ArgumentException(msg);
        }
    }
}
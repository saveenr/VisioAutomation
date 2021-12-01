namespace VisioAutomation.Internal
{
    internal readonly partial struct VisioObjectTarget
    {
        public string[] GetFormulas(
            ShapeSheet.Streams.StreamArray stream)
        {
            if (stream.Array.Length == 0)
            {
                return new string[0];
            }

            System.Array formulas_sa = null;

            if (this.Category == Internal.VisioObjectCategory.Shape)
            {
                this.Shape.GetFormulasU(stream.Array, out formulas_sa);
            }
            else if (this.Category == Internal.VisioObjectCategory.Master)
            {
                this.Master.GetFormulasU(stream.Array, out formulas_sa);
            }
            else if (this.Category == Internal.VisioObjectCategory.Page)
            {
                this.Page.GetFormulasU(stream.Array, out formulas_sa);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            var formulas = Internal.ShapesheetHelpers.SystemArrayToTypedArray<string>(formulas_sa);
            return formulas;
        }

        public TResult[] GetResults<TResult>(
            ShapeSheet.Streams.StreamArray stream,
            object[] unitcodes)
        {
            if (stream.Array.Length == 0)
            {
                return new TResult[0];
            }

            Internal.ShapesheetHelpers.EnforceValid_ResultType(typeof(TResult));

            var flags = Internal.ShapesheetHelpers.GetVisGetSetArgsFromType(typeof(TResult));
            System.Array results_sa = null;

            if (this.Category == Internal.VisioObjectCategory.Shape)
            {
                this.Shape.GetResults(stream.Array, (short) flags, unitcodes, out results_sa);
            }
            else if (this.Category == Internal.VisioObjectCategory.Master)
            {
                this.Master.GetResults(stream.Array, (short) flags, unitcodes, out results_sa);
            }
            else if (this.Category == Internal.VisioObjectCategory.Page)
            {
                this.Page.GetResults(stream.Array, (short) flags, unitcodes, out results_sa);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            var results = Internal.ShapesheetHelpers.SystemArrayToTypedArray<TResult>(results_sa);
            return results;
        }

        public int SetFormulas(
            ShapeSheet.Streams.StreamArray stream,
            object[] formulas, short flags)
        {
            Internal.ShapesheetHelpers.ValidateStreamLengthFormulas(stream, formulas);

            int val = 0;
            if (this.Category == VisioObjectCategory.Shape)
            {
                val = this.Shape.SetFormulas(stream.Array, formulas, flags);
            }
            else if (this.Category == VisioObjectCategory.Master)
            {
                val = this.Master.SetFormulas(stream.Array, formulas, flags);
            }
            else if (this.Category == VisioObjectCategory.Page)
            {
                val = this.Page.SetFormulas(stream.Array, formulas, flags);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            return val;
        }

        public int SetResults(
            ShapeSheet.Streams.StreamArray stream, object[] unitcodes, object[] results, short flags)
        {
            Internal.ShapesheetHelpers.ValidateStreamLengthResults(stream, results);

            int val = 0;
            if (this.Category == VisioObjectCategory.Shape)
            {
                val = this.Shape.SetResults(stream.Array, unitcodes, results, flags);
            }
            else if (this.Category == VisioObjectCategory.Master)
            {
                val = this.Master.SetResults(stream.Array, unitcodes, results, flags);
            }
            else if (this.Category == VisioObjectCategory.Page)
            {
                val = this.Page.SetResults(stream.Array, unitcodes, results, flags);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            return val;
        }
    }
}
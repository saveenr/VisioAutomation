namespace VisioAutomation.ShapeSheet.Writers
{
    public class WriterBase
    {
        public bool BlastGuards { get; set; }
        public bool TestCircular { get; set; }

        protected Microsoft.Office.Interop.Visio.VisGetSetArgs ComputeGetResultFlags()
        {
            var flags = this.combine_blastguards_and_testcircular_flags();

            flags |= Microsoft.Office.Interop.Visio.VisGetSetArgs.visGetStrings;

            return flags;
        }

        protected Microsoft.Office.Interop.Visio.VisGetSetArgs ComputeGetFormulaFlags()
        {
            var common_flags = this.combine_blastguards_and_testcircular_flags();
            var formula_flags = (short)Microsoft.Office.Interop.Visio.VisGetSetArgs.visSetUniversalSyntax;
            var combined_flags = (short)common_flags | formula_flags;
            return (Microsoft.Office.Interop.Visio.VisGetSetArgs)combined_flags;
        }

        private Microsoft.Office.Interop.Visio.VisGetSetArgs combine_blastguards_and_testcircular_flags()
        {
            var f_bg = this.BlastGuards ? Microsoft.Office.Interop.Visio.VisGetSetArgs.visSetBlastGuards : 0;
            var f_tc = this.TestCircular ? Microsoft.Office.Interop.Visio.VisGetSetArgs.visSetTestCircular : 0;

            var flags = ((short)f_bg) | ((short)f_tc);
            return (Microsoft.Office.Interop.Visio.VisGetSetArgs)flags;
        }

    }
}
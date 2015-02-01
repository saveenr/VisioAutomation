namespace TestVisioAutomation
{
    public class DiagnosticDebugContext : VisioAutomation.Scripting.Context
    {
        public DiagnosticDebugContext()
        {
        }

        public override void WriteDebug(string s)
        {
            string msg = string.Format("DEBUG: {0}", s);
            this.writeline(msg);
        }

        public override void WriteError(string s)
        {
            string msg = string.Format("ERROR: {0}", s);
            this.writeline(s);
        }

        public override void WriteUser(string s)
        {
            string msg = string.Format("USER: {0}", s);
            this.writeline(s);
        }

        public override void WriteVerbose(string s)
        {
            string msg = string.Format("VERBOSE: {0}", s);
            this.writeline(s);
        }

        public override void WriteWarning(string s)
        {
            string msg = string.Format("WARNING: {0}", s);
            this.writeline(s);
        }

        private void writeline(string s)
        {
            System.Diagnostics.Debug.WriteLine(s);
        }
    }
}
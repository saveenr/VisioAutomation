using System.Diagnostics;
using VASCRIPT=VisioAutomation.Scripting;

namespace TestVisioAutomation
{
    public class DiagnosticDebugContext : VASCRIPT.Context
    {
        public DiagnosticDebugContext()
        {
        }

        public override void WriteDebug(string s)
        {
            string msg = $"DEBUG: {s}";
            this.writeline(msg);
        }

        public override void WriteError(string s)
        {
            string msg = $"ERROR: {s}";
            this.writeline(s);
        }

        public override void WriteUser(string s)
        {
            string msg = $"USER: {s}";
            this.writeline(s);
        }

        public override void WriteVerbose(string s)
        {
            string msg = $"VERBOSE: {s}";
            this.writeline(s);
        }

        public override void WriteWarning(string s)
        {
            string msg = $"WARNING: {s}";
            this.writeline(s);
        }

        private void writeline(string s)
        {
            Debug.WriteLine(s);
        }
    }
}
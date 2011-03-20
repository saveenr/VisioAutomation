namespace VisioIPy
{
    public class VisioIPySessionOptions : VisioAutomation.Scripting.SessionOptions
    {
        private VisioIPySession vi;

        public VisioIPySessionOptions(VisioIPySession vi)
        {
            this.vi=vi;
        }

        public override void WriteDebug(string s)
        {
            if (this.vi.Debug)
            {
                System.Console.WriteLine("DEBUG: {0}", s);                
            }
        }

        public override void WriteError(string s)
        {
            System.Console.WriteLine("ERROR: {0}", s);
        }

        public override void WriteUser(string s)
        {
            System.Console.WriteLine(s);
        }

        public override void WriteVerbose(string s)
        {
            if (this.vi.Verbose)
            {
                System.Console.WriteLine("VERBOSE: {0}", s);
            }
        }
    }
}
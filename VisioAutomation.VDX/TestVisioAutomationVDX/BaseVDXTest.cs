namespace TestVisioAutomationVDX
{
    public class BaseVDXTest
    {
        protected string GetTestResultsOutPath(string path)
        {
            return System.IO.Path.Combine(this.TestResultsOutFolder, path);
        }

        private static string tr_out_folder;

        protected string TestResultsOutFolder
        {
            get
            {
                if (tr_out_folder == null)
                {
                    var asm = System.Reflection.Assembly.GetExecutingAssembly();
                    tr_out_folder = System.IO.Path.GetDirectoryName(asm.Location);
                }
                return tr_out_folder;
            }
        }
    }
}
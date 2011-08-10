using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomationSamples
{
    public class Program
    {
        private static void Main(string[] args)
        {
            var form = new FormSampleRunner();
            form.ShowDialog();
        }
    }
}
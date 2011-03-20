using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Linq;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        private void writelineconsole(string s)
        {
            System.Console.WriteLine(s);
        }

        public void WriteLine(string s)
        {
            this.writelineconsole(s);
        }

        public void WriteLine(string fmt, params object[] items)
        {
            this.writelineconsole(string.Format(fmt, items));
        }

        public void Print(object o)
        {
            this.WriteLine(o.ToString());
        }
    }
}
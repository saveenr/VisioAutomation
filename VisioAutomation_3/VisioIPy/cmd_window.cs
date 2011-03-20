using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public void SetApplicationWindowSize(int w, int h)
        {
            this.ScriptingSession.SetApplicationWindowSize(w, h);
        }

        public void SetApplicationWindowSize(System.Drawing.Size size)
        {
            this.ScriptingSession.SetApplicationWindowSize(size.Width, size.Height);
        }

        public System.Drawing.Size GetApplicationWindowSize()
        {
            var size = this.ScriptingSession.GetApplicationWindowSize();
            return size;
        }
    }
}
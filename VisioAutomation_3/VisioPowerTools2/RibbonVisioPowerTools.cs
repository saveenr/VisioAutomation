using Microsoft.Office.Tools.Ribbon;
using MOC = Microsoft.Office.Core;
using VA = VisioAutomation;

namespace VisioPowerTools2
{
    public partial class RibbonVisioPowerTools
    {
        private void RibbonVisioPowerTools_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var fildef = new VA.Effects.TwoColorGlow();
            fildef.TopColor = new VA.Drawing.ColorRGB(0xff0000);
            fildef.BottomColor = new VA.Drawing.ColorRGB(0xff00f0);
            fildef.TopTransparency = 0.0;
            fildef.BottomTransparency = 1.0;
            fildef.Scale = 1.4;

            var fmt = fildef.GetFormat();

            ThisAddIn.ScriptingSession.Format.SetFormat(fmt);
        }
    }
}

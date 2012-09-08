using System.Diagnostics;
using VA = VisioAutomation;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.XamlTune
{
    public static class XamlTuneHelper
    {
        public static void ExportSelectionAsXAML(Session scripting_session, string filename)
        {
            VA.ExportHelper.ExportSelectionAsXAML2(scripting_session.Selection.Get(), filename, s=>scripting_session.Output.Write(OutputStream.Verbose,s));
        }
    }
}
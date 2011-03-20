using System.Linq;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting
{
    internal static class MiscScriptingUtil
    {
        internal static Dictionary<VA.Drawing.AlignmentHorizontal, IVisio.VisUICmds> halign_to_cmd = new Dictionary
            <VA.Drawing.AlignmentHorizontal, IVisio.VisUICmds>
                    {
                        { VA.Drawing.AlignmentHorizontal.Left, IVisio.VisUICmds.visCmdAlignObjectLeft},
                        { VA.Drawing.AlignmentHorizontal.Center, IVisio.VisUICmds.visCmdAlignObjectCenter },
                        { VA.Drawing.AlignmentHorizontal.Right, IVisio.VisUICmds.visCmdAlignObjectRight }
                    };

        internal static Dictionary<VA.Drawing.AlignmentVertical, IVisio.VisUICmds> valign_to_cmd = new Dictionary
            <VA.Drawing.AlignmentVertical, IVisio.VisUICmds>
                    {
                        { VA.Drawing.AlignmentVertical.Top, IVisio.VisUICmds.visCmdAlignObjectTop },
                        { VA.Drawing.AlignmentVertical.Center, IVisio.VisUICmds.visCmdAlignObjectMiddle },
                        { VA.Drawing.AlignmentVertical.Bottom, IVisio.VisUICmds.visCmdAlignObjectBottom }
                    };
    }
}
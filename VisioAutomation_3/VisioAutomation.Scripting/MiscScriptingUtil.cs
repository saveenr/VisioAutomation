using System;
using System.Linq;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting
{
    internal static class MiscScriptingUtil
    {
        public static IVisio.VisUICmds AlignmentToUICmd(VA.Drawing.AlignmentHorizontal a)
        {
            if (a==VA.Drawing.AlignmentHorizontal.Left)
            {
                return IVisio.VisUICmds.visCmdAlignObjectLeft;
            }
            if (a==VA.Drawing.AlignmentHorizontal.Center)
            {
                return IVisio.VisUICmds.visCmdAlignObjectCenter;
            }
            if (a == VA.Drawing.AlignmentHorizontal.Right)
            {
                return IVisio.VisUICmds.visCmdAlignObjectRight;
            }
            else
            {
                throw new ArgumentOutOfRangeException();
            }
        }

        public static IVisio.VisUICmds AlignmentToUICmd(VA.Drawing.AlignmentVertical a)
        {
            if (a == VA.Drawing.AlignmentVertical.Top) { return IVisio.VisUICmds.visCmdAlignObjectTop; }
            if (a==VA.Drawing.AlignmentVertical.Center) {   return IVisio.VisUICmds.visCmdAlignObjectMiddle; }
            if (a == VA.Drawing.AlignmentVertical.Bottom) { return IVisio.VisUICmds.visCmdAlignObjectBottom; }
            else
            {
                throw new ArgumentOutOfRangeException();
            }
        }
    }
}
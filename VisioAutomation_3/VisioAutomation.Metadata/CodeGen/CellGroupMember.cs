using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VA=VisioAutomation;

namespace VisioAutomation.Metadata.CodeGen
{
    public class CellGroupMember
    {
        public string MemberName;
        public VA.Metadata.Cell Cell;
        public string DataTypeName;
    }
}

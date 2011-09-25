using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VA=VisioAutomation;

namespace VisioAutomation.Metadata.CodeGen
{
    public class VAGenCode
    {
        public static string GetCode()
        {
            var md = VA.Metadata.MetadataDB.Load();
            var cg = new VA.Metadata.CodeGen.CellGroup("ShapeFormatCells");
            cg.Add("FillForegnd", md.GetCellByNameCode("FillForegnd"));

            string code = cg.GenCode();
            return code;
        }
    }
}

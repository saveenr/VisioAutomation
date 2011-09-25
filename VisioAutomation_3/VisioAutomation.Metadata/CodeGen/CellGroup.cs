using System;
using System.Collections.Generic;
using System.Linq;
using VA=VisioAutomation;

namespace VisioAutomation.Metadata.CodeGen
{
    public class CellGroup
    {
        public string Name;
        public string Parent;
        public List<VA.Metadata.CodeGen.CellGroupMember> Cells;

        public CellGroup(string name)
        {
            this.Name=name;
            this.Parent = "X";
            this.Cells = new List<CellGroupMember>();
        }

        public string GenCode()
        {
            var sb = new System.Text.StringBuilder();
            this.Start(sb);
            this.Constructor(sb);
            this.End(sb);

            return sb.ToString();
        }
        //        public VA.ShapeSheet.CellData<int> FillBkgnd { get; set; }

        private void Start(System.Text.StringBuilder sb)
        {
            sb.AppendFormat("public class {0} : {1}\n", this.Name, this.Parent);
            sb.AppendFormat("{{\n\n");
            foreach (var cell in this.Cells)
            {
                sb.AppendFormat("VA.ShapeSheet.CellData<{0}> {1};","double",cell.MemberName);
            }
        }

        private void End(System.Text.StringBuilder sb)
        {
            sb.AppendFormat("}}\n\n");
        }

        private void Constructor(System.Text.StringBuilder sb)
        {
            sb.AppendFormat("public class {0} : {1}\n", this.Name, this.Parent);
            foreach (var cell in this.Cells)
            {
                sb.AppendFormat("{0};",cell.MemberName);
            }
        }

        public void Add(string name, VA.Metadata.Cell cell)
        {
            var m = new VA.Metadata.CodeGen.CellGroupMember();
            m.MemberName = name;
            m.Cell = cell;
        }
    }
}

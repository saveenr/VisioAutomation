using System;
using System.Collections.Generic;
using System.Linq;
using VA=VisioAutomation;

namespace VisioAutomation.CodeGen
{
    public class CellGroup
    {
        public string Name;
        public string Parent;
        public Type DataType;
        public bool ForSection;
        public List<VA.CodeGen.CellGroupMember> Cells;

        public CellGroup(string name)
        {
            this.Name=name;

            this.Cells = new List<CellGroupMember>();
            this.ForSection = false;
        }

        public string GenCode()
        {
            this.Parent = this.ForSection ? "CellSectionDataGroup" : "CellDataGroup";
            var sb = new System.Text.StringBuilder();
            var cs = new VA.CodeGen.CSharpWriter(sb);

            this.Start(cs);
            cs.WriteLine();
            this.ApplyFunc(cs);
            cs.WriteLine();
            this.CellsFromRow(cs);
            cs.WriteLine();

            string rt_a;
            string rt_b;


            if (this.ForSection)
            {
                rt_a = string.Format("IList< List<{0}>>", this.Name);
                rt_b = string.Format("List<{0}>", this.Name);
            }
            else
            {
                rt_a = string.Format("IList<{0}>", this.Name);
                rt_b = string.Format("{0}", this.Name);
            }
            
            cs.WriteLine("internal static {0} GetCells(IVisio.Page page, IList<int> shapeids)", rt_a);
            cs.StartBlock();
            cs.WriteLine("var query = new ShapeFormatQuery();");
            cs.WriteLine("return {0}._GetCells(page, shapeids, query, get_cells_from_row);", this.Parent);
            cs.EndBlock();

            cs.WriteLine("internal static {0} GetCells(IVisio.Shape shape)", rt_b);
            cs.StartBlock();
            cs.WriteLine("var query = new ShapeFormatQuery();");
            cs.WriteLine("return {0}._GetCells(page, shapeids, query, get_cells_from_row);", this.Parent);
            cs.EndBlock();

            this.Query(cs);
            this.End(cs);

            return sb.ToString();
        }
        //        public VA.ShapeSheet.CellData<int> FillBkgnd { get; set; }

        private void Start(VA.CodeGen.CSharpWriter sb)
        {
            sb.WriteLine("public class {0} : {1}", this.Name, this.Parent);
            sb.StartBlock();
            foreach (var cell in this.Cells)
            {
                sb.WriteLine("public VA.ShapeSheet.CellData<{0}> {1} {{get;set;}}",cell.DataTypeName,cell.MemberName);
            }
        }

        private void End(VA.CodeGen.CSharpWriter sb)
        {
            sb.EndBlock();
        }

        private void ApplyFunc(VA.CodeGen.CSharpWriter sb)
        {
            if (this.ForSection)
            {
                sb.WriteLine("protected override void _Apply(VA.ShapeSheet.CellDataGroup.ApplyFormula func, short row)");
                
            }
            else
            {
                sb.WriteLine("protected override void _Apply(VA.ShapeSheet.CellDataGroup.ApplyFormula func)");
            }
            sb.StartBlock();
            foreach (var cell in this.Cells)
            {
                if (this.ForSection)
                {
                    sb.WriteLine("func(ShapeSheet.SRCConstants.{0}.ForRow(row), this.{1}.Formula);", cell.Cell, cell.MemberName);
                    
                }
                else
                {
                    sb.WriteLine("func(ShapeSheet.SRCConstants.{0}, this.{1}.Formula);", cell.Cell, cell.MemberName);
                }
            }
            sb.EndBlock();
        }

        private void CellsFromRow(VA.CodeGen.CSharpWriter sb)
        {
            sb.WriteLine("private static ShapeFormatCells get_cells_from_row(ShapeFormatQuery query, VA.ShapeSheet.Query.QueryDataSet<double> qds, int row)");
            sb.StartBlock();
            sb.WriteLine("var cells = new {0}();;", this.Name);
            foreach (var cell in this.Cells)
            {
                string to = "To"+cell.DataTypeName.Substring(0, 1).ToUpper() + cell.DataTypeName.Substring(1);
                sb.WriteLine("cells.{0}= qds.GetItem(row, query.{0}).{1}();", cell.MemberName, to);
            }
            sb.EndBlock();
        }

        public void Add(string name, string cell, string datatype)
        {
            var m = new VA.CodeGen.CellGroupMember();
            m.MemberName = name;
            m.Cell = cell;
            m.DataTypeName = datatype;
            this.Cells.Add(m);
        }

        public void Add(string cell, string datatype)
        {
            this.Add(cell,cell,datatype);
        }

        private void Query(VA.CodeGen.CSharpWriter csw)
        {
            csw.WriteLine();
            csw.WriteLine();
            string Queryname = this.Name + "Query";
            csw.WriteLine("class {0}", Queryname);
            csw.StartBlock();
            foreach (var cell in this.Cells)
            {
                csw.WriteLine("public VA.ShapeSheet.Query.CellQueryColumn {0} {{ get; set; }};", cell.MemberName);
            }
            csw.WriteLine();
            csw.WriteLine("public {0}()", Queryname);
            csw.StartBlock();
            foreach (var cell in this.Cells)
            {
                csw.WriteLine("this.{0} = this.AddColumn(VA.ShapeSheet.SRCConstants.{1}, \"{2}\" );", cell.MemberName,
                                cell.Cell, cell.MemberName);

            }
            csw.EndBlock();
            csw.EndBlock();
        }

    }
}

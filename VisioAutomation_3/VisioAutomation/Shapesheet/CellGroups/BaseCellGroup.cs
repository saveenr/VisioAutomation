using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class BaseCellGroup
    {
        // Delegates
        protected delegate void ApplyFormula(VA.ShapeSheet.SRC src, VA.ShapeSheet.FormulaLiteral formula);
        
        public class CellMember
        {
            public System.Type DataType { get; private set; }
            public string Name { get; private set; }

            public CellMember(System.Type datatype, string name)
            {
                this.DataType = datatype;
                this.Name = name;
            }

            public override string ToString()
            {
                return string.Format("{0}.{1}", this.Name, this.DataType);
            }
        }

        public List<CellMember> GetCellMembers()
        {
            var t = this.GetType();
            var bingingflags = System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.GetProperty;
            var members = t.GetProperties(bingingflags);
            var targets = members.Where(m => IsCellData(m));

            var items = new List<CellMember>();
            foreach (var target in targets)
            {
                var ga = target.PropertyType.GetGenericArguments();
                var celldata_datatype = ga[0];
                var cm = new CellMember(celldata_datatype, target.Name);
                items.Add(cm);
            }

            return items;
        }

        private bool IsCellData( System.Reflection.PropertyInfo p)
        {
            return ((p.PropertyType == typeof (VA.ShapeSheet.CellData<int>))
                || (p.PropertyType == typeof(VA.ShapeSheet.CellData<double>))
                || (p.PropertyType == typeof(VA.ShapeSheet.CellData<string>))
                || (p.PropertyType == typeof(VA.ShapeSheet.CellData<bool>)));
        }

        protected static IEnumerable<VA.ShapeSheet.Data.QueryDataRow<T>> EnumRows<T>(VA.ShapeSheet.Data.QueryDataSet<T> qds)
        {
            for (int row = 0; row < qds.Rows.Count; row++)
            {
                yield return qds.GetRow(row);
            }
        }
    }
}
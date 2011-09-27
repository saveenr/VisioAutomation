using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet
{
    public abstract class CellDataGroup
    {
        // Delegates
        protected delegate void ApplyFormula(VA.ShapeSheet.SRC src, VA.ShapeSheet.FormulaLiteral formula);
        protected delegate TObj RowToObject<TObj, TQuery>(TQuery query, VA.ShapeSheet.Data.QueryDataRow<double> qdr) where TQuery : VA.ShapeSheet.Query.CellQuery;
        
        protected abstract void _Apply(ApplyFormula func);

        public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short shapeid)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(shapeid, src, f));
        }

        public void Apply(VA.ShapeSheet.Update.SRCUpdate update)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(src, f));
        }

        protected static IList<TCells> _GetObjectsFromRows<TCells, TQuery>(IVisio.Page page, IList<int> shapeids, TQuery q, RowToObject<TCells, TQuery> row_to_cells_func) where TQuery : VA.ShapeSheet.Query.CellQuery
        {
            var qds = q.GetFormulasAndResults<double>(page, shapeids);
            var obj_list = new List<TCells>(qds.RowCount);
            foreach (var row in qds.EnumRows())
            {
                var obj = row_to_cells_func(q, row);
                obj_list.Add(obj);                
            }

            return obj_list;
        }

        protected static TCells _GetObjectFromSingleRow<TCells, TQuery>(IVisio.Shape shape, TQuery query, RowToObject<TCells, TQuery> row_to_obj_func) where TQuery : VA.ShapeSheet.Query.CellQuery
        {
            var qds = query.GetFormulasAndResults<double>(shape);
            var qdr = qds.GetRow(0);
            return row_to_obj_func(query, qdr);
        }

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
    }
}
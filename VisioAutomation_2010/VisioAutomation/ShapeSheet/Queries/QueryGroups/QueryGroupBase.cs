using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Queries.QueryGroups
{
    public abstract class QueryGroupBase
    {
        protected SRCFormulaPair newpair(ShapeSheet.SRC src, ShapeSheet.FormulaLiteral formula)
        {
            return new SRCFormulaPair(src, formula);
        }

        public abstract IEnumerable<SRCFormulaPair> Pairs { get; }
    }
}
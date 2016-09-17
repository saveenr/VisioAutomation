using VisioAutomation.ShapeSheet.Queries;

namespace VisioAutomation.ShapeSheet.CellGroups.Queries
{
    public abstract class CellGroupQuery<TCellGroup, TResult>
    {
        protected VisioAutomation.ShapeSheet.Queries.Query query;

        protected CellGroupQuery()
        {
            this.query = new Query();
        }

        protected abstract void validate_query();

        public abstract TCellGroup CellDataToCellGroup(ShapeSheet.CellData<TResult>[] row);

    }
}
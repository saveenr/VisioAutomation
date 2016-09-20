using VisioAutomation.ShapeSheet.Queries;

namespace VisioAutomation.ShapeSheet.CellGroups.Queries
{
    public abstract class CellGroupQuery<TCellGroup>
    {
        protected VisioAutomation.ShapeSheet.Queries.Query query;

        protected CellGroupQuery()
        {
            this.query = new Query();
        }

        protected abstract void validate_query();

        public abstract TCellGroup CellDataToCellGroup(ShapeSheet.CellData[] row);

    }
}
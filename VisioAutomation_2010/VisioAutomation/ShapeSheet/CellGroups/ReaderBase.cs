using VisioAutomation.ShapeSheet.Queries;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class ReaderBase<TCellGroup>
    {
        protected VisioAutomation.ShapeSheet.Queries.Query query;

        protected ReaderBase()
        {
            this.query = new Query();
        }

        protected abstract void validate_query();

        public abstract TCellGroup CellDataToCellGroup(ShapeSheet.CellData[] row);

    }
}
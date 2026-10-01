namespace VisioAutomation.Models.Data
{
    public class DataTableModel
    {
        public DataTableModel()
        {
            // The cell size used before CellWidth and CellHeight were honored.
            this.CellWidth = 1.0;
            this.CellHeight = 1.0;
        }

        public System.Data.DataTable DataTable { get; set; }
        public double CellWidth { get; set; }
        public double CellHeight { get; set; }
        public double CellSpacing { get; set; }
    }
}
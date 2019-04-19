namespace VisioScripting.Models
{
    public enum ZoomToObject
    {
        Page,
        PageWidth,
        Selection
    }

    public class DataTableModel
    {

        public System.Data.DataTable DataTable { get; set; }

        public double CellWidth { get; set; }

        public double CellHeight { get; set; }

        public double CellSpacing { get; set; }
    }

}
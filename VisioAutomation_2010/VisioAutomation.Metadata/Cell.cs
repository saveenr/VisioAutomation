namespace VisioAutomation.Metadata
{
    public class Cell
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public string NameCode { get; set; }
        public string NameFormatString { get; set; }
        public string Object { get; set; }
        public string NameType { get; set; }
        public string DataType { get; set; }
        public string ContentType { get; set; }
        public string Unit { get; set; }
        public string SectionIndex { get; set; }
        public string RowIndex { get; set; }
        public string MinVersion { get; set; }
        public string MaxVersion { get; set; }
        public string CellIndex { get; set; }
        public string MSDN { get; set; }

        public override string ToString()
        {
            return this.NameCode;
        }
    }
}
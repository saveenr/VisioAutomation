namespace ExcelUtil
{
    public class RowData
    {
        public string[] Value { get; private set; }
        public string[] Type { get; private set; }
        public ColumnDefinition[] Column { get; private set; }

        public RowData(int capacity)
        {
            this.Value = new string[capacity];
            this.Type = new string[capacity];
            this.Column = new ColumnDefinition[capacity];
        }

        public int Length
        {
            get { return this.Value.Length; }
        }

        public void Clear()
        {
            for (int i = 0; i < this.Length; i++)
            {
                this.Value[i] = null;
                this.Type[i] = null;
                this.Column[i] = null;
            }
        }
    }
}

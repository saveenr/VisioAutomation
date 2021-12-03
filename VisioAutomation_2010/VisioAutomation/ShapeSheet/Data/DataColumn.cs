namespace VisioAutomation.ShapeSheet.Data
{
    public class DataColumn
    {
        public string Name { get; protected set; }
        public int Ordinal { get; protected set; }
        public Core.Src Src { get; }

        internal DataColumn(int ordinal, string name, Core.Src src)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new System.ArgumentException("name");
            }

            this.Src = src;
            this.Name = name;
            this.Ordinal = ordinal;
        }

        protected DataColumn(int ordinal, string name)
        {
        }

        public static implicit operator int(DataColumn col)
        {
            return col.Ordinal;
        }
    }
}
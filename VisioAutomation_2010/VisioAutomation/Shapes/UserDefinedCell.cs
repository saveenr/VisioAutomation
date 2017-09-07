namespace VisioAutomation.Shapes
{
    public class UserDefinedCell
    {
        public string Name { get; set; }
        public UserDefinedCellCells Cells { get; set; }

        public UserDefinedCell(string name)
        {
            UserDefinedCellHelper.CheckValidName(name);
            this.Name = name;
            this.Cells = new UserDefinedCellCells();
        }

        public UserDefinedCell(string name, string value)
        {
            UserDefinedCellHelper.CheckValidName(name);

            if (value == null)
            {
                throw new System.ArgumentNullException(nameof(value));
            }

            this.Name = name;
            this.Cells = new UserDefinedCellCells();
            this.Cells.Value = value;
        }

        public UserDefinedCell(string name, string value, string prompt)
        {
            UserDefinedCellHelper.CheckValidName(name);

            if (value == null)
            {
                throw new System.ArgumentNullException(nameof(value));
            }

            this.Name = name;
            this.Cells.Value = value;
            this.Cells.Prompt = prompt;
        }

        public UserDefinedCell(string name, UserDefinedCellCells cells)
        {
            UserDefinedCellHelper.CheckValidName(name);

            if (cells == null)
            {
                throw new System.ArgumentNullException(nameof(cells));
            }

            this.Name = name;
            this.Cells = cells;
        }


        public override string ToString()
        {
            string s = string.Format("(Name={0},Value={1},Prompt={2})", this.Name, this.Cells.Value, this.Cells.Prompt);
            return s;
        }

    }
}
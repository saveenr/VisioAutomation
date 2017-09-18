namespace VisioScripting.Models
{
    public class UserDefinedCell
    {
        public string Name { get; set; }
        public VisioAutomation.Shapes.UserDefinedCellCells Cells { get; set; }

        public UserDefinedCell(string name)
        {
            VisioAutomation.Shapes.UserDefinedCellHelper.CheckValidName(name);
            this.Name = name;
            this.Cells = new VisioAutomation.Shapes.UserDefinedCellCells();
        }

        public UserDefinedCell(string name, string value)
        {
            VisioAutomation.Shapes.UserDefinedCellHelper.CheckValidName(name);

            if (value == null)
            {
                throw new System.ArgumentNullException(nameof(value));
            }

            this.Name = name;
            this.Cells = new VisioAutomation.Shapes.UserDefinedCellCells();
            this.Cells.Value = value;
        }

        public UserDefinedCell(string name, string value, string prompt)
        {
            VisioAutomation.Shapes.UserDefinedCellHelper.CheckValidName(name);

            if (value == null)
            {
                throw new System.ArgumentNullException(nameof(value));
            }

            this.Name = name;
            this.Cells.Value = value;
            this.Cells.Prompt = prompt;
        }

        public UserDefinedCell(string name, VisioAutomation.Shapes.UserDefinedCellCells cells)
        {
            VisioAutomation.Shapes.UserDefinedCellHelper.CheckValidName(name);

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
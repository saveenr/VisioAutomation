namespace VisioAutomation.DOM
{
    public class Hyperlink
    {
        public string Name { get; set; }
        public string Address { get; set; }

        public Hyperlink(string name, string address)
        {
            this.Name = name;
            this.Address = address;
        }
    }
}
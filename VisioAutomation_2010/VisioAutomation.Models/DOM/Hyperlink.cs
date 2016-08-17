namespace VisioAutomation.Models.DOM
{
	public class Hyperlink
	{
	    public string Name { get; set; }
	    public string Description { get; set; }
	    public string Address { get; set; }
	    public string SubAddress { get; set; }
	    public string ExtraInfo { get; set; }
	    public string Frame { get; set; }
	    public string SortKey { get; set; }
	    public bool NewWindow { get; set; }
	    public bool Default { get; set; }
	    public bool Invisible { get; set; }
	 	 
	    public Hyperlink(string name, string address)
	    {
	        this.Name = name;
	        this.Address = address;
	    }
	}
}


namespace VisioAutomation.Models.Dom;

public class Connector : Shape
{
    public BaseShape From { get; }
    public BaseShape To { get; }
        
    public Connector(BaseShape from, BaseShape to, IVisio.Master master) :
        base(master,-3,-3)
    {
        this.Master = new MasterRef(master);
        this.From = from;
        this.To = to;
    }

    public Connector(BaseShape from, BaseShape to, string mastername, string stencilname) :
        base(mastername,stencilname, new VisioAutomation.Geometry.Point(-3,-3) )
    {
        this.Master = new MasterRef(mastername, stencilname);
        this.From = from;
        this.To = to;
    }
}
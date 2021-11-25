namespace VisioAutomation.ShapeSheet.Query;

public class SectionQueryColumns : Columns
{
    public IVisio.VisSectionIndices SectionIndex { get; }

    internal SectionQueryColumns(IVisio.VisSectionIndices section)
    {
        this.SectionIndex = section;
    }
}
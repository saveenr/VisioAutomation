namespace VisioAutomation.Scripting.Models
{
    [System.Flags]
    public enum FormatPaintCategory
    {
        Fill = 0x01 << 0 ,
        Character = 0x01 << 1,
        Line = 0x01 << 2,
        Shadow = 0x01 << 3,
        Paragraph = 0x01 << 4 
    }
}
namespace VisioAutomation.Scripting
{
    [System.Flags]
    public enum FormatCategory
    {
        Fill = 0x01 << 0 ,
        Character = 0x01 << 1,
        Line = 0x01 << 2,
        Shadow = 0x01 << 3,
        Paragraph = 0x01 << 4 
    }
}
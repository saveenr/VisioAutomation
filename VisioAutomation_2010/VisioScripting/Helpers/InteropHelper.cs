
namespace VisioScripting.Helpers;

public static class InteropHelper
{
    private static bool _static_initialized = false;
    private static Dictionary<string, Models.EnumType> _static_g_name_to_enum;
    private static List<System.Type> _static_g_types;

    private static void _initialize()
    {
        if (!InteropHelper._static_initialized)
        {
            InteropHelper._static_g_types = typeof(IVisio.Application).Assembly.GetExportedTypes()
                .Where(t => t.IsPublic)
                .Where(t => !t.Name.StartsWith("tag"))
                .OrderBy(t => t.Name)
                .ToList();
            InteropHelper._static_g_name_to_enum = InteropHelper._static_g_types
                .Where(t => t.IsEnum)
                .Select(i => new Models.EnumType(i))
                .ToDictionary(i => i.Name, i => i);
            InteropHelper._static_initialized = true;
        }
    }

    public static List<Models.EnumType> GetEnums()
    {
        InteropHelper._initialize();
        return InteropHelper._static_g_name_to_enum.Values.ToList();
    }

    public static Models.EnumType GetEnum(string name)
    {
        InteropHelper._initialize();
        return InteropHelper._static_g_name_to_enum[name];
    }

    public static Models.EnumType GetEnum(System.Type t)
    {
        InteropHelper._initialize();
        return InteropHelper._static_g_name_to_enum[t.Name];
    }
}
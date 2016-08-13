namespace VisioAutomation.Extensions
{
    public static class ApplicationMethods
    {
        public static void Quit(this Microsoft.Office.Interop.Visio.Application app, bool force_close)
        {
            Application.ApplicationHelper.Quit(app,force_close);
        }
    }
}
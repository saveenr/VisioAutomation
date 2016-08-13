namespace VisioAutomation.Scripting.Commands
{
    public class Command
    {
        public System.Reflection.MethodInfo MethodInfo;

        public Command(System.Reflection.MethodInfo mi)
        {
            this.MethodInfo = mi;
        }
    }
}
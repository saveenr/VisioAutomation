using System.Collections.Generic;

namespace VisioScripting.Commands
{
    public class Command
    {
        protected readonly System.Reflection.MethodInfo MethodInfo;

        public string Name => this.MethodInfo.Name;
        public System.Type Type => this.MethodInfo.ReturnType;
        public readonly string TypeDisplayName;

        internal Command(System.Reflection.MethodInfo methodinfo)
        {
            this.MethodInfo = methodinfo;
            this.TypeDisplayName = VisioScripting.Helpers.ReflectionHelper.GetNiceTypeName(this.MethodInfo.ReturnType);
        }

        public IEnumerable<CommandParameter> GetParameters()
        {
            var method_params = this.MethodInfo.GetParameters();
            foreach (var methodparam in method_params)
            {
                string cmdparam_typedispname = VisioScripting.Helpers.ReflectionHelper.GetNiceTypeName(methodparam.ParameterType);
                string cmdparam_name = methodparam.Name;
                System.Type cmdtype = methodparam.ParameterType;
                var cmdparam = new CommandParameter(cmdparam_name, cmdtype, cmdparam_typedispname);

                yield return cmdparam;
            }
        }
    }
}
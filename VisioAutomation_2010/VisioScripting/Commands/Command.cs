using System.Collections.Generic;

namespace VisioScripting.Commands
{
    public class Command
    {
        private readonly System.Reflection.MethodInfo MethodInfo;

        public string Name => this.MethodInfo.Name;
        public System.Type ReturnType => this.MethodInfo.ReturnType;
        public readonly string ReturnTypeDisplayName;
        public bool ReturnsValue => this.ReturnType != typeof(void);

        internal Command(System.Reflection.MethodInfo methodinfo)
        {
            this.MethodInfo = methodinfo;
            this.ReturnTypeDisplayName = VisioScripting.Helpers.ReflectionHelper.GetNiceTypeName(this.MethodInfo.ReturnType);
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
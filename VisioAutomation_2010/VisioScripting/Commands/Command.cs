using System.Collections.Generic;

namespace VisioScripting.Commands;

public class Command
{
    private readonly System.Reflection.MethodInfo _method_info;

    public string Name => this._method_info.Name;
    public System.Type ReturnType => this._method_info.ReturnType;
    public readonly string ReturnTypeDisplayName;
    public bool ReturnsValue => this.ReturnType != typeof(void);

    internal Command(System.Reflection.MethodInfo methodinfo)
    {
        this._method_info = methodinfo;
        this.ReturnTypeDisplayName = VisioScripting.Helpers.ReflectionHelper.GetNiceTypeName(this._method_info.ReturnType);
    }

    public IEnumerable<CommandParameter> GetParameters()
    {
        var method_params = this._method_info.GetParameters();
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
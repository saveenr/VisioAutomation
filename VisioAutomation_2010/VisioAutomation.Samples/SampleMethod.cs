namespace VisioAutomationSamples;

public class SampleMethod
{
    public string Name;
    public System.Reflection.MethodInfo Method;

    public void Run()
    {
        this.Method.Invoke(null, null);
    }
}
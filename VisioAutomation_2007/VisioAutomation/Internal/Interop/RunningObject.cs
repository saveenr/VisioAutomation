namespace VisioAutomation.Internal.Interop
{
    internal class RunningObject
    {
        private readonly string _display_name;
        private readonly object _object;
        private readonly System.Guid _classid;

        public RunningObject(string name, object obj, System.Guid classid)
        {
            this._display_name = name;
            this._object = obj;
            this._classid = classid;
        }

        public string DisplayName
        {
            get { return _display_name; }
        }

        public object Object
        {
            get { return _object; }
        }

        public System.Guid ClassId
        {
            get { return _classid; }
        }
    }
}
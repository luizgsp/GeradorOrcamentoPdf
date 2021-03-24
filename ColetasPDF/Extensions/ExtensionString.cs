namespace System
{
    static class ExtensionString
    {
        public static string Completing(this string obj, int count, char character)
        {

            if (obj.Length == count) { return obj; }
            if (obj.Length > count) { return obj.Substring(0, count); }
            for (int i = obj.Length; i <= count; i++)
            {
                obj += character;
            }
            return obj;
        }
    }
}

using System.Collections.Generic;

namespace VisioAutomation.Models.InternalTree
{
    static class DictionaryUtil
    {
        /// <summary>
        /// If the key is in the dictionary, returns the value associated with the key.
        /// If the key is not in the dictionary, returns the default value (the key will not be added the the dictionary)
        /// </summary>
        /// <typeparam name="K"></typeparam>
        /// <typeparam name="V"></typeparam>
        /// <param name="dic"></param>
        /// <param name="key"></param>
        /// <param name="defval"></param>
        /// <returns></returns>
        public static V GetValue<K, V>(Dictionary<K, V> dic, K key, V defval)
        {
            V the_item;
            bool contains = dic.TryGetValue(key, out the_item);

            if (!contains)
            {
                return defval;
            }

            return the_item;
        }
    }
}
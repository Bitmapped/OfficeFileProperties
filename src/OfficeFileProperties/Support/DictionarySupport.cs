using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeFileProperties.Support
{
    static public class DictionarySupport
    {
        static public string Serialize(this IDictionary<string, object> dictionary)
        {
            // Generate string.
            var propertyString = String.Empty;

            foreach (var item in dictionary)
            {
                // Insert break if needed.
                if (propertyString != string.Empty)
                {
                    propertyString += " ||| ";
                }

                // Include in try-catch block in case of problem with converting value to string.
                try
                {
                    // Add new item onto string.
                    propertyString += item.Key + " ::: " + item.Value.ToString();
                }
                catch { }
            }

            return propertyString;
        }
    }
}

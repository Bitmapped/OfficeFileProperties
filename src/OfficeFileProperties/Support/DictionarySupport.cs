using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeFileProperties.Support
{
    static public class DictionarySupport
    {
        static public string Serialize(this Dictionary<string, string> dictionary)
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

                // Add new item onto string.
                propertyString += item.Key + " ::: " + item.Value;
            }

            return propertyString;
        }
    }
}

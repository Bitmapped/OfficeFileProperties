using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeFileProperties.Exceptions
{
    public class NullPropertyException : System.Exception
    {
        public NullPropertyException()
            : base()
        {

        }

        public NullPropertyException(string message)
            : base(message)
        {

        }

        public NullPropertyException(string message, Exception inner)
            : base(message, inner)
        {

        }
    }
}

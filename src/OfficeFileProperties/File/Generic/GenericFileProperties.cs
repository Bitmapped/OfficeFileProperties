using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeFileProperties.File.Generic
{
    /// <summary>
    /// File properties for a generic file.
    /// </summary>
    class GenericFileProperties : FileProperties
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public GenericFileProperties()
        {
            // Store file type.
            this.fileType = FileTypeEnum.OtherType;
        }        
    }
}

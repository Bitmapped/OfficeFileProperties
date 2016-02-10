using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeFileProperties.File
{
    
    /// <summary>
    /// Properties specific to all files.
    /// </summary>
    public interface IFileProperties
    {
        FileTypeEnum FileType { get; }

        DateTime CreatedTimeLocal { get; }
        DateTime CreatedTimeUtc { get; }
        DateTime ModifiedTimeLocal { get; }
        DateTime ModifiedTimeUtc { get; }
        string Filename { get; }
    }
}

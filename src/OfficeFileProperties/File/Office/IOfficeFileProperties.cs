using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeFileProperties.File;

namespace OfficeFileProperties.File.Office
{
    /// <summary>
    /// Additional properties specific to Microsoft Office files.
    /// </summary>
    public interface IOfficeFileProperties: IFileProperties
    {
        string Author { get; }
        string Company { get; }
        string Title { get; }
        SortedList<string, string> CustomProperties { get; }
        string CustomPropertiesString { get; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeFileProperties
{
    /// <summary>
    /// Enum for tracking the file type.
    /// </summary>
    public enum FileTypeEnum {
        Unknown = 0,
        MicrosoftWord = 1,
        MicrosoftExcel = 2,
        MicrosoftAccess = 4,
        MicrosoftPowerPoint = 8,
        OtherType = 16
    };
}

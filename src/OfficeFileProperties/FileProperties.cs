using System;
using System.Collections.Generic;
using OfficeFileProperties.Support;

namespace OfficeFileProperties
{
    public class FileProperties
    {

        #region Properties

        /// <summary>
        /// Author name
        /// </summary>
        public string Author { get; internal set; }

        /// <summary>
        /// Company name
        /// </summary>
        public string Company { get; internal set; }

        /// <summary>
        /// Created Date in local time
        /// </summary>
        public DateTime? CreatedDateLocal
        {
            get
            {
                if (CreatedDateUtc.HasValue)
                {
                    return CreatedDateUtc.Value.ToLocalTime();
                }
                else
                {
                    return null;
                }
            }
            internal set
            {
                if (value.HasValue)
                {
                    CreatedDateUtc = value.Value.ToUniversalTime();
                }
                else
                {
                    CreatedDateUtc = null;
                }
            }
        }

        /// <summary>
        /// Created Date in UTC time
        /// </summary>
        public DateTime? CreatedDateUtc { get; internal set; }

        /// <summary>
        /// Custom Properties
        /// </summary>
        public Dictionary<string, string> CustomProperties { get; internal set; }

        /// <summary>
        /// Serialize Custom Properties as a string.
        /// </summary>
        public string CustomPropertiesString
        {
            get
            {
                if (CustomProperties != null)
                {
                    return CustomProperties.Serialize();
                }
                else
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Filename
        /// </summary>
        public string Filename { get; internal set; }

        /// <summary>
        /// Type of file
        /// </summary>
        public FileTypeEnum FileType { get; internal set; }

        /// <summary>
        /// Modified Date in local time
        /// </summary>
        public DateTime? ModifiedDateLocal
        {
            get
            {
                if (ModifiedDateUtc.HasValue)
                {
                    return ModifiedDateUtc.Value.ToLocalTime();
                }
                else
                {
                    return null;
                }
            }
            internal set
            {
                if (value.HasValue)
                {
                    ModifiedDateUtc = value.Value.ToUniversalTime();
                }
                else
                {
                    ModifiedDateUtc = null;
                }
            }
        }

        /// <summary>
        /// Modified Date in UTC time
        /// </summary>
        public DateTime? ModifiedDateUtc { get; internal set; }

        /// <summary>
        /// Title
        /// </summary>
        public string Title { get; internal set; }

        #endregion Properties
    }
}
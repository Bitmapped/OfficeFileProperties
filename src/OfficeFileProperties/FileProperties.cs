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
        /// Created Time in local time
        /// </summary>
        public DateTime? CreatedTimeLocal
        {
            get
            {
                if (CreatedTimeUtc.HasValue)
                {
                    return CreatedTimeUtc.Value.ToLocalTime();
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
                    CreatedTimeUtc = value.Value.ToUniversalTime();
                }
                else
                {
                    CreatedTimeUtc = null;
                }
            }
        }

        /// <summary>
        /// Created Time in UTC time
        /// </summary>
        public DateTime? CreatedTimeUtc { get; internal set; }

        /// <summary>
        /// Custom Properties
        /// </summary>
        public IDictionary<string, object> CustomProperties { get; internal set; }

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
        /// Modified Time in local time
        /// </summary>
        public DateTime? ModifiedTimeLocal
        {
            get
            {
                if (ModifiedTimeUtc.HasValue)
                {
                    return ModifiedTimeUtc.Value.ToLocalTime();
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
                    ModifiedTimeUtc = value.Value.ToUniversalTime();
                }
                else
                {
                    ModifiedTimeUtc = null;
                }
            }
        }

        /// <summary>
        /// Modified Time in UTC time
        /// </summary>
        public DateTime? ModifiedTimeUtc { get; internal set; }

        /// <summary>
        /// Title
        /// </summary>
        public string Title { get; internal set; }

        #endregion Properties
    }
}
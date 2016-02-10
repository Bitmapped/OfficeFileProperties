using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeFileProperties.File
{
    /// <summary>
    /// File properties for a generic file.
    /// </summary>
    class FileProperties : IFileProperties
    {

        // Define internal variables for storing file properties.
        internal string filename;
        internal bool fileLoaded;
        internal DateTime createdTimeUtc, modifiedTimeUtc;
        internal FileTypeEnum fileType = FileTypeEnum.Unknown;

        /// <summary>
        /// Access file type.
        /// </summary>
        public FileTypeEnum FileType
        {
            get
            {
                return this.fileType;
            }
        }

        

        /// <summary>
        /// Filename
        /// </summary>
        public string Filename
        {
            get
            {
                // Check that file has been loaded.
                if (!this.fileLoaded)
                {
                    throw new InvalidOperationException("No file has been loaded.");
                }

                return this.filename;
            }
        }

        /// <summary>
        /// File creation time in UTC.
        /// </summary>
        public DateTime CreatedTimeUtc
        {
            get
            {
                // Check that file has been loaded.
                if (!this.fileLoaded)
                {
                    throw new InvalidOperationException("No file has been loaded.");
                }

                return this.createdTimeUtc;
            }
        }

        /// <summary>
        /// File creation time in local time.
        /// </summary>
        public DateTime CreatedTimeLocal
        {
            get
            {
                return this.CreatedTimeUtc.ToLocalTime();
            }
        }

        /// <summary>
        /// File modification time in UTC.
        /// </summary>
        public DateTime ModifiedTimeUtc
        {
            get
            {
                // Check that file has been loaded.
                if (!this.fileLoaded)
                {
                    throw new InvalidOperationException("No file has been loaded.");
                }

                return this.modifiedTimeUtc;
            }
        }

        /// <summary>
        /// File modification time in local time.
        /// </summary>
        public DateTime ModifiedTimeLocal
        {
            get
            {
                return this.ModifiedTimeUtc.ToLocalTime();
            }
        }
    }
}

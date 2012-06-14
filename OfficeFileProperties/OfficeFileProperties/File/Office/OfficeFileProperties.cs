using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeFileProperties.File.Office
{
    /// <summary>
    /// File properties for an Office file.
    /// </summary>
    class OfficeFileProperties : FileProperties, IOfficeFileProperties
    {
        // Define private variables.
        internal string author, company, title;
        internal SortedList<string, string> customProperties;

        /// <summary>
        /// Constructor
        /// </summary>
        public OfficeFileProperties()
        {
            // Create empty sorted list for custom properties.
            this.customProperties = new SortedList<string, string>();
        }

        /// <summary>
        /// Title
        /// </summary>
        public string Title
        {
            get
            {
                // Check that file has been loaded.
                if (!this.fileLoaded)
                {
                    throw new InvalidOperationException("No file has been loaded.");
                }

                return this.title;
            }
        }

        /// <summary>
        /// Company
        /// </summary>
        public string Company
        {
            get
            {
                // Check that file has been loaded.
                if (!this.fileLoaded)
                {
                    throw new InvalidOperationException("No file has been loaded.");
                }

                return this.company;
            }
        }

        /// <summary>
        /// Author
        /// </summary>
        public string Author
        {
            get
            {
                // Check that file has been loaded.
                if (!this.fileLoaded)
                {
                    throw new InvalidOperationException("No file has been loaded.");
                }

                return this.author;
            }
        }

        /// <summary>
        /// CustomProperties
        /// </summary>
        public SortedList<string, string> CustomProperties
        {
            get
            {
                // Check that file has been loaded.
                if (!this.fileLoaded)
                {
                    throw new InvalidOperationException("No file has been loaded.");
                }

                return this.customProperties;
            }
        }
    }
}

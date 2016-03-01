using System;
using System.Collections.Generic;

namespace OfficeFileProperties
{
    /// <summary>
    /// Interface for working with files.
    /// </summary>
    public interface IFile
    {
        #region Properties

        /// <summary>
        /// Author name
        /// </summary>
        string Author { get; set; }

        /// <summary>
        /// Company anme
        /// </summary>
        string Company { get; set; }

        /// <summary>
        /// Created Time in local time
        /// </summary>
        DateTime? CreatedTimeLocal { get; set; }

        /// <summary>
        /// Created Time in UTC time
        /// </summary>
        DateTime? CreatedTimeUtc { get; set; }

        /// <summary>
        /// Custom Properties
        /// </summary>
        IDictionary<string, object> CustomProperties { get; }

        /// <summary>
        /// Serialize Custom Properties as a string.
        /// </summary>
        string CustomPropertiesString { get; }

        /// <summary>
        /// Filename
        /// </summary>
        string Filename { get; }

        /// <summary>
        /// Type of file
        /// </summary>
        FileTypeEnum FileType { get; }

        /// <summary>
        /// Indicator if the file is currently open
        /// </summary>
        bool IsOpen { get; }

        /// <summary>
        /// Modified Time in local time
        /// </summary>
        DateTime? ModifiedTimeLocal { get; set; }

        /// <summary>
        /// Modified Time in UTC time
        /// </summary>
        DateTime? ModifiedTimeUtc { get; set; }

        /// <summary>
        /// Title
        /// </summary>
        string Title { get; set; }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Closes file.
        /// </summary>
        void CloseFile();

        /// <summary>
        /// Gets FileProperties object loaded with properties for current file.
        /// </summary>
        /// <returns>Loaded FileProperties object</returns>
        FileProperties GetFileProperties();

        /// <summary>
        /// Opens file.
        /// </summary>
        void OpenFile();

        #endregion Methods
    }
}
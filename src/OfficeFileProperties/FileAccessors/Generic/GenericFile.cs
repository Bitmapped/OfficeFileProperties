using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeFileProperties.FileAccessors.Generic
{
    /// <summary>
    /// Class for using generic files.
    /// </summary>
    public class GenericFile : FileBase<FileInfo>
    {
        #region Constructors

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="filename">Filename to open.</param>
        public GenericFile(string filename) : base(filename)
        { }

        #endregion Constructors

        /// <summary>
        /// Type of file.
        /// </summary>
        public override FileTypeEnum FileType => FileTypeEnum.OtherType;

        #region Methods

        /// <summary>
        /// Closes file.
        /// </summary>
        /// <param name="saveChanges"></param>
        public override void CloseFile(bool saveChanges = false)
        {
            // Clear file object.
            this.File = null;
        }

        public override bool IsWritable => false;

        public override bool IsOpen => this.File != null;

        public override bool IsReadable => this.File != null;

        /// <summary>
        /// Opens file.
        /// </summary>
        /// <param name="writable"></param>
        public override void OpenFile(bool writable = false)
        {
            // Open file.
            this.File = new FileInfo(this.Filename);
        }

        /// <summary>
        /// Created date in UTC time
        /// </summary>
        public override DateTime? CreatedTimeUtc => this.File.CreationTimeUtc;

        /// <summary>
        /// Created date in UTC time
        /// </summary>
        public override DateTime? ModifiedTimeUtc => this.File.LastWriteTimeUtc;

        /// <summary>
        /// Author name
        /// </summary>
        public override string Author => null;

        /// <summary>
        /// Company name
        /// </summary>
        public override string Company => null;

        /// <summary>
        /// Comments (description)
        /// </summary>
        public override string Comments => null;

        /// <summary>
        /// Title
        /// </summary>
        public override string Title => null;

        /// <summary>
        /// Custom Properties
        /// </summary>
        public override IDictionary<string, object> CustomProperties => null;

        #endregion Methods
    }
}
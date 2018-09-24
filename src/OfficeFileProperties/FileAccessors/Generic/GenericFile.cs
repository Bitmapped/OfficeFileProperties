using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
        public override FileTypeEnum FileType
        {
            get
            {
                return FileTypeEnum.OtherType;
            }
        }

        #region Methods

        /// <summary>
        /// Closes file.
        /// </summary>
        /// <param name="saveChanges"></param>
        public override void CloseFile(bool saveChanges = false)
        {
            // Mark file as closed.
            this.IsOpen = false;

            // Clear file object.
            this.File = null;
        }

        public override bool IsWritable
        {
            get { return false; }
        }

        /// <summary>
        /// Opens file.
        /// </summary>
        /// <param name="writable"></param>
        public override void OpenFile(bool writable = false)
        {
            // Open file.
            this.File = new FileInfo(this.Filename);

            // Mark file as open.
            this.IsOpen = true;
        }

        /// <summary>
        /// Created date in UTC time
        /// </summary>
        public override DateTime? CreatedTimeUtc
        {
            get
            {
                return this.File.CreationTimeUtc;
            }
        }

        /// <summary>
        /// Created date in UTC time
        /// </summary>
        public override DateTime? ModifiedTimeUtc
        {
            get
            {
                return this.File.LastWriteTimeUtc;
            }
        }

        /// <summary>
        /// Author name
        /// </summary>
        public override string Author
        {
            get
            {
                return null;
            }
        }

        /// <summary>
        /// Company name
        /// </summary>
        public override string Company
        {
            get
            {
                return null;
            }
        }

        /// <summary>
        /// Comments (description)
        /// </summary>
        public override string Comments
        {
            get
            {
                return null;
            }
        }

        /// <summary>
        /// Title
        /// </summary>
        public override string Title
        {
            get
            {
                return null;
            }
        }

        /// <summary>
        /// Custom Properties
        /// </summary>
        public override IDictionary<string, object> CustomProperties
        {
            get
            {
                return null;
            }
        }

        #endregion Methods
    }
}

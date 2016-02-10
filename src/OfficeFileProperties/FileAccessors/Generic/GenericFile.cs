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
        public override void CloseFile()
        {
            // Mark file as closed.
            this.IsOpen = false;

            // Clear file object.
            this.FileAccessor = null;
        }

        /// <summary>
        /// Opens file.
        /// </summary>
        public override void OpenFile()
        {
            // Open file.
            this.FileAccessor = new FileInfo(this.Filename);

            // Mark file as open.
            this.IsOpen = true;
        }

        /// <summary>
        /// Created date in UTC time
        /// </summary>
        public override DateTime? CreatedDateUtc
        {
            get
            {
                return this.FileAccessor.CreationTimeUtc;
            }
        }

        /// <summary>
        /// Created date in UTC time
        /// </summary>
        public override DateTime? ModifiedDateUtc
        {
            get
            {
                return this.FileAccessor.LastWriteTimeUtc;
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
        public override Dictionary<string, string> CustomProperties
        {
            get
            {
                return null;
            }
        }

        #endregion Methods
    }
}

using DSOFile;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeFileProperties.FileAccessors.Dso
{
    /// <summary>
    /// Class for using Microsoft Word DOCX files.
    /// </summary>
    public class DsoFile : FileBase<OleDocumentProperties>
    {

        #region Constructors

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="filename">Filename to open.</param>
        public DsoFile(string filename) : base(filename)
        { }
        #endregion Constructors

        #region Properties

        /// <summary>
        /// Author name
        /// </summary>
        public override string Author
        {
            get
            {
                // Ensure file is open.
                if (!this.IsOpen)
                {
                    throw new InvalidOperationException("File is not open.");
                }

                return this.File.SummaryProperties.Author;
            }
        }

        /// <summary>
        /// Company name
        /// </summary>
        public override string Company
        {
            get
            {
                // Ensure file is open.
                if (!this.IsOpen)
                {
                    throw new InvalidOperationException("File is not open.");
                }

                return this.File.SummaryProperties.Company;
            }
        }

        /// <summary>
        /// Created date in UTC time
        /// </summary>
        public override DateTime? CreatedTimeUtc
        {
            get
            {
                if (this.File.SummaryProperties.DateCreated is DateTime)
                {
                    return ((DateTime)this.File.SummaryProperties.DateCreated).ToUniversalTime();
                }
                else
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Custom Properties
        /// </summary>
        public override IDictionary<string, string> CustomProperties
        {
            get
            {
                // Ensure file is open.
                if (!this.IsOpen)
                {
                    throw new InvalidOperationException("File is not open.");
                }

                if (this.File.CustomProperties == null)
                {
                    return new Dictionary<string, string>();
                }

                // Iterate manually and specify types because of dynamic COM object.
                var customProperties = new Dictionary<string, string>();
                foreach (CustomProperty item in this.File.CustomProperties)
                {
                    customProperties.Add(item.ToString(), item.get_Value().ToString());
                }

                return customProperties;
            }
        }

        /// <summary>
        /// Type of file.
        /// </summary>
        public override FileTypeEnum FileType
        {
            get
            {
                switch (new FileInfo(this.Filename).Extension)
                {
                    case ".xls":
                        return FileTypeEnum.MicrosoftExcel;

                    case ".ppt":
                        return FileTypeEnum.MicrosoftPowerPoint;

                    case ".doc":
                        return FileTypeEnum.MicrosoftWord;

                    default:
                        return FileTypeEnum.OtherType;
                }
            }
        }

        /// <summary>
        /// Created date in UTC time
        /// </summary>
        public override DateTime? ModifiedTimeUtc
        {
            get
            {
                if (this.File.SummaryProperties.DateLastSaved is DateTime)
                {
                    return ((DateTime)this.File.SummaryProperties.DateLastSaved).ToUniversalTime();
                }
                else
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Title
        /// </summary>
        public override string Title
        {
            get
            {
                // Ensure file is open.
                if (!this.IsOpen)
                {
                    throw new InvalidOperationException("File is not open.");
                }

                return this.File.SummaryProperties.Title;
            }
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Closes file.
        /// </summary>
        public override void CloseFile()
        {
            // Mark file as closed.
            this.IsOpen = false;

            // Close file if it still is accessible.
            if (this.File != null)
            {
                // Close file.
                this.File.Close();
            }

            // Clear file object.
            this.File = null;
        }

        /// <summary>
        /// Opens file.
        /// </summary>
        public override void OpenFile()
        {
            // Open file.
            this.File = new OleDocumentProperties();
            this.File.Open(Filename, true);

            // Mark file as open.
            this.IsOpen = true;
        }

        #endregion Methods
    }
}

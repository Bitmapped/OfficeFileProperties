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

        #region Methods

        /// <summary>
        /// Closes file.
        /// </summary>
        public override void CloseFile()
        {
            // Mark file as closed.
            this.IsOpen = false;

            // Close file.
            this.FileAccessor.Close();

            // Clear file object.
            this.FileAccessor = null;
        }

        /// <summary>
        /// Opens file.
        /// </summary>
        public override void OpenFile()
        {
            // Open file.
            this.FileAccessor = new OleDocumentProperties();
            this.FileAccessor.Open(Filename, true);

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
                if (this.FileAccessor.SummaryProperties.DateCreated is DateTime)
                {
                    return ((DateTime)this.FileAccessor.SummaryProperties.DateCreated).ToUniversalTime();
                }
                else
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Created date in UTC time
        /// </summary>
        public override DateTime? ModifiedDateUtc
        {
            get
            {
                if(this.FileAccessor.SummaryProperties.DateLastSaved is DateTime)
                {
                    return ((DateTime)this.FileAccessor.SummaryProperties.DateLastSaved).ToUniversalTime();
                }
                else
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Author name
        /// </summary>
        public override string Author
        {
            get
            {
                return this.FileAccessor.SummaryProperties.Author;
            }
        }

        /// <summary>
        /// Company name
        /// </summary>
        public override string Company
        {
            get
            {
                return this.FileAccessor.SummaryProperties.Company;
            }
        }

        /// <summary>
        /// Title
        /// </summary>
        public override string Title
        {
            get
            {
                return this.FileAccessor.SummaryProperties.Title;
            }
        }

        /// <summary>
        /// Custom Properties
        /// </summary>
        public override Dictionary<string, string> CustomProperties
        {
            get
            {
                if (this.FileAccessor.CustomProperties == null)
                {
                    return new Dictionary<string, string>();
                }

                // Iterate manually and specify types because of dynamic COM object.
                var customProperties = new Dictionary<string, string>();
                foreach (CustomProperty item in this.FileAccessor.CustomProperties)
                {
                    customProperties.Add(item.ToString(), item.get_Value().ToString());
                }

                return customProperties;
            }
        }

        #endregion Methods
    }
}

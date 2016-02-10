using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.CustomProperties;

namespace OfficeFileProperties.FileAccessors.OpenXml
{
    /// <summary>
    /// Class for using Microsoft Excel XLSX files.
    /// </summary>
    public class XlsxFile : FileBase<SpreadsheetDocument>
    {
        #region Constructors

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="filename">Filename to open.</param>
        public XlsxFile(string filename) : base(filename)
        { }
        #endregion Constructors

        /// <summary>
        /// Type of file.
        /// </summary>
        public override FileTypeEnum FileType
        {
            get
            {
                return FileTypeEnum.MicrosoftExcel;
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

            // Dispose of file.
            this.FileAccessor.Dispose();

            // Clear file object.
            this.FileAccessor = null;
        }

        /// <summary>
        /// Opens file.
        /// </summary>
        public override void OpenFile()
        {
            // Open file.
            this.FileAccessor = SpreadsheetDocument.Open(this.Filename, false);

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
                if (this.FileAccessor.PackageProperties.Created.HasValue)
                {
                    return this.FileAccessor.PackageProperties.Created.Value.ToUniversalTime();
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
                if (this.FileAccessor.PackageProperties.Modified.HasValue)
                {
                    return this.FileAccessor.PackageProperties.Modified.Value.ToUniversalTime();
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
                return this.FileAccessor.PackageProperties.Creator;
            }
        }

        /// <summary>
        /// Company name
        /// </summary>
        public override string Company
        {
            get
            {
                return (this.FileAccessor.ExtendedFilePropertiesPart.Properties.Company != null) ? this.FileAccessor.ExtendedFilePropertiesPart.Properties.Company.InnerText : null;
            }
        }

        /// <summary>
        /// Title
        /// </summary>
        public override string Title
        {
            get
            {
                return this.FileAccessor.PackageProperties.Title;
            }
        }

        /// <summary>
        /// Custom Properties
        /// </summary>
        public override Dictionary<string, string> CustomProperties
        {
            get
            {
                if (this.FileAccessor.CustomFilePropertiesPart == null)
                {
                    return new Dictionary<string, string>();
                }

                var customProperties = this.FileAccessor.CustomFilePropertiesPart.Properties
                                            .Select(p => (CustomDocumentProperty)p)
                                            .ToDictionary(cp => cp.Name.Value, cp => cp.InnerText.ToString());

                return customProperties;
            }
        }

        #endregion Methods
    }
}

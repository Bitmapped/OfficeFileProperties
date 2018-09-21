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
    /// Class for using Microsoft PowerPoint PPTX files.
    /// </summary>
    public class PptxFile : FileBase<PresentationDocument>
    {
        #region Constructors

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="filename">Filename to open.</param>
        public PptxFile(string filename) : base(filename)
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

            // Close file if it still is accessible.
            if (this.File != null)
            {
                // Close file.
                this.File.Close();

                // Dispose of file.
                this.File.Dispose();
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
            this.File = PresentationDocument.Open(this.Filename, false);

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
                // Ensure file is open.
                if (!this.IsOpen)
                {
                    throw new InvalidOperationException("File is not open.");
                }

                if (this.File.PackageProperties.Created.HasValue)
                {
                    return this.File.PackageProperties.Created.Value.ToUniversalTime();
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
        public override DateTime? ModifiedTimeUtc
        {
            get
            {
                // Ensure file is open.
                if (!this.IsOpen)
                {
                    throw new InvalidOperationException("File is not open.");
                }

                if (this.File.PackageProperties.Modified.HasValue)
                {
                    return this.File.PackageProperties.Modified.Value.ToUniversalTime();
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
                // Ensure file is open.
                if (!this.IsOpen)
                {
                    throw new InvalidOperationException("File is not open.");
                }

                return this.File.PackageProperties.Creator;
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

                return ((this.File.ExtendedFilePropertiesPart != null) && (this.File.ExtendedFilePropertiesPart.Properties.Company != null)) ? this.File.ExtendedFilePropertiesPart.Properties.Company.InnerText : null;
            }
        }

        /// <summary>
        /// Comments (description)
        /// </summary>
        public override string Comments
        {
            get
            {
                // Ensure file is open.
                if (!this.IsOpen)
                {
                    throw new InvalidOperationException("File is not open.");
                }

                return this.File.PackageProperties.Description;
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

                return this.File.PackageProperties.Title;
            }
        }

        /// <summary>
        /// Custom Properties
        /// </summary>
        public override IDictionary<string, object> CustomProperties
        {
            get
            {
                // Ensure file is open.
                if (!this.IsOpen)
                {
                    throw new InvalidOperationException("File is not open.");
                }

                if (this.File.CustomFilePropertiesPart == null)
                {
                    return new Dictionary<string, object>();
                }

                var customProperties = this.File.CustomFilePropertiesPart.Properties
                                            .Select(p => (CustomDocumentProperty)p)
                                            .ToDictionary<CustomDocumentProperty, string, object>(cp => cp.Name.Value, cp => cp.InnerText);

                return customProperties;
            }
        }

        #endregion Methods
    }
}

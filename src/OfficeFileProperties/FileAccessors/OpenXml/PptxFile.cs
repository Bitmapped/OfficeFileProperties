using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using Properties = DocumentFormat.OpenXml.ExtendedProperties.Properties;

namespace OfficeFileProperties.FileAccessors.OpenXml
{
    /// <summary>
    /// Class for using Microsoft PowerPoint PPTX files.
    /// </summary>
    public class PptxFile : OpenXmlFileBase<PresentationDocument>
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
        public override FileTypeEnum FileType => FileTypeEnum.MicrosoftExcel;

        #region Methods

        /// <summary>
        /// Opens file.
        /// </summary>
        /// <param name="writable"></param>
        public override void OpenFile(bool writable = false)
        {
            // Open file.
            this.File = PresentationDocument.Open(this.Filename, writable);
        }

        /// <summary>
        /// Company name
        /// </summary>
        public override string Company
        {
            get
            {
                // Ensure file is open.
                this.TestFileOpen();

                return this.File.ExtendedFilePropertiesPart?.Properties?.Company?.InnerText;
            }
            set
            {
                // Ensure file is writable.
                this.TestFileWritable();

                // Add extended file properties part if it does not exist.
                if (this.File.ExtendedFilePropertiesPart == null)
                {
                    this.File.AddExtendedFilePropertiesPart();
                }

                // Add properties part if it does not exist.
                if (this.File.ExtendedFilePropertiesPart.Properties == null)
                {
                    this.File.ExtendedFilePropertiesPart.Properties = new Properties();
                }

                this.File.ExtendedFilePropertiesPart.Properties.Company = new Company(value);
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
                                           .Select(p => (CustomDocumentProperty) p)
                                           .ToDictionary<CustomDocumentProperty, string, object>(cp => cp.Name.Value, cp => cp.InnerText);

                return customProperties;
            }
            set => throw new NotImplementedException();
        }

        #endregion Methods
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.ExtendedProperties;

namespace OfficeFileProperties.File.Office.OpenXml
{
    /// <summary>
    /// Obtain properties of Microsoft Excel XLSX files.
    /// </summary>
    class XlsxFile : IFile
    {
        // Define private variables.
        private string filename;
        private SpreadsheetDocument file;
        private OfficeFileProperties fileProperties;

        /// <summary>
        /// Access file properties.
        /// </summary>
        public IFileProperties FileProperties
        {
            get
            {
                // Get file if loaded.
                if ((this.fileProperties != null) && (this.fileProperties.fileLoaded))
                {
                    return this.fileProperties;
                }
                else
                {
                    // File has not been loaded.
                    throw new InvalidOperationException("No file has been loaded.");
                }
            }
        }        

        /// <summary>
        /// Constructor
        /// </summary>
        public XlsxFile()
        {
            ClearProperties();
        }

        /// <summary>
        /// Constructor, also loads specified file.
        /// </summary>
        /// <param name="filename">File to load.</param>
        public XlsxFile(string filename) : this()
        {
            // Load specified file.
            this.LoadFile(filename);
        }

        /// <summary>
        /// Forget about loaded file.
        /// </summary>
        private void CloseFile()
        {
            // Close file.
            this.file.Close();
            this.file.Dispose();

            // Clear file object.
            this.file = null;
        }

        /// <summary>
        /// Clears values of loaded properties.
        /// </summary>
        private void ClearProperties()
        {
            // Clear loaded file, properties object.
            this.filename = null;
            this.fileProperties = null;
        }

        /// <summary>
        /// Loads requested file, saves its properties.
        /// </summary>
        /// <param name="filename"></param>
        public void LoadFile(string filename)
        {
            // Clear loaded properties.
            ClearProperties();

            // Store filename.
            this.filename = filename;

            // Load file.
            this.file = SpreadsheetDocument.Open(filename, false);

            // Loads file properties.
            LoadProperties();

            // Since file cannot be written to, close it immediately.
            CloseFile();
        }

        /// <summary>
        /// Load values from file into private variables.
        /// </summary>
        private void LoadProperties()
        {
            // If file hasn't been loaded, throw exception.
            if (this.file == null)
            {
                throw new InvalidOperationException("No file is currently loaded.");
            }

            // Create new file properties object.
            this.fileProperties = new OfficeFileProperties();

            // filename
            this.fileProperties.filename = this.filename;

            // filetype
            this.fileProperties.fileType = OfficeFileProperties.FileTypeEnum.MicrosoftExcel;
            
            // createdTimeUtc
            try
            {
                this.fileProperties.createdTimeUtc = this.file.PackageProperties.Created.Value.ToUniversalTime();
            }
            catch
            {
                this.fileProperties.createdTimeUtc = new DateTime(1, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            }

            // modifiedTimeUtc
            // Try getting actual time, otherwise return dummy value if it fails.
            try
            {
                this.fileProperties.modifiedTimeUtc = this.file.PackageProperties.Modified.Value.ToUniversalTime();
            }
            catch
            {
                this.fileProperties.modifiedTimeUtc = new DateTime(1, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            }

            // author
            this.fileProperties.author = this.file.PackageProperties.Creator;

            // title
            this.fileProperties.title = this.file.PackageProperties.Title;

            // company
            this.fileProperties.company = this.file.ExtendedFilePropertiesPart.Properties.Company.InnerText;

            // Load custom properties.
            if (this.file.CustomFilePropertiesPart.Properties != null)
            {
                // Use Linq to get listing of properties.
                var customProperties = this.file.CustomFilePropertiesPart.Properties
                             .Select(p => (CustomDocumentProperty)p);
                
                // Iterate through custom properties.
                foreach (var cp in customProperties)
                {
                    this.fileProperties.customProperties.Add(cp.Name.Value, cp.InnerText.ToString());
                }
            }

            // Mark properties as loaded.
            this.fileProperties.fileLoaded = true;
        }

    }
}

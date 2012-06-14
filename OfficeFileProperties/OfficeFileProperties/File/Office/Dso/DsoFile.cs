using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using DSOFile;

namespace OfficeFileProperties.File.Office.Dso
{
    /// <summary>
    /// Obtain properties of Microsoft Office 97-2003 files.
    /// </summary>
    class DsoFile : IFile
    {
        // Define private variables.
        private string filename;
        private OleDocumentProperties file;
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
        public DsoFile()
        {
            ClearProperties();
        }

        /// <summary>
        /// Constructor, also loads specified file.
        /// </summary>
        /// <param name="filename">File to load.</param>
        public DsoFile(string filename)
            : this()
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
            this.file = new OleDocumentProperties();
            this.file.Open(filename, true);

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
            var fileInfo = new FileInfo(this.filename);
            switch (fileInfo.Extension)
            {
                case ".xls":
                    this.fileProperties.fileType = FileTypeEnum.MicrosoftExcel;
                    break;

                case ".ppt":
                    this.fileProperties.fileType = FileTypeEnum.MicrosoftPowerPoint;
                    break;

                case ".doc":
                    this.fileProperties.fileType = FileTypeEnum.MicrosoftWord;
                    break;

                default:
                    this.fileProperties.fileType = FileTypeEnum.OtherType;
                    break;
            }

            // createdTimeUtc
            // Try getting actual time, otherwise return dummy value if it fails.
            try
            {
                this.fileProperties.createdTimeUtc = ((DateTime)this.file.SummaryProperties.DateCreated).ToUniversalTime();
            }
            catch
            {
                this.fileProperties.createdTimeUtc = new DateTime(1, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            }

            // modifiedTimeUtc
            // Try getting actual time, otherwise return dummy value if it fails.
            try
            {
                this.fileProperties.modifiedTimeUtc = ((DateTime)this.file.SummaryProperties.DateLastSaved).ToUniversalTime();
            }
            catch
            {
                this.fileProperties.modifiedTimeUtc = new DateTime(1, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            }

            // author
            this.fileProperties.author = this.file.SummaryProperties.Author;

            // title
            this.fileProperties.title = this.file.SummaryProperties.Title;

            // company
            this.fileProperties.company = this.file.SummaryProperties.Company;

            // Load custom properties.
            foreach (CustomProperty cp in this.file.CustomProperties)
            {
                this.fileProperties.customProperties.Add(cp.Name.ToString(), cp.get_Value().ToString());
            }

            // Mark properties as loaded.
            this.fileProperties.fileLoaded = true;
        }

    }
}

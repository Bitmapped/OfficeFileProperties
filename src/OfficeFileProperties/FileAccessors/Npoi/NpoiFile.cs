using NPOI.HPSF;
using NPOI.HPSF.Extractor;
using NPOI.POIFS.FileSystem;
using OfficeFileProperties.Support.Npoi;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;

namespace OfficeFileProperties.FileAccessors.Npoi
{
    /// <summary>
    /// Class for using Microsoft Office 97-2003 files.
    /// </summary>
    public class NpoiFile : FileBase<NpoiPropertiesOnlyDocument>
    {

        /// <summary>
        /// Stream for NPOI file
        /// </summary>
        private FileStream _fileStream;

        #region Constructors

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="filename">Filename to open.</param>
        public NpoiFile(string filename) : base(filename)
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

                return this.SummaryInformation.Author;
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

                return this.DocumentSummaryInformation.Company;
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

                return this.SummaryInformation.Comments;
            }
        }

        /// <summary>
        /// Created date in UTC time
        /// </summary>
        public override DateTime? CreatedTimeUtc
        {
            get
            {
                if (this.SummaryInformation.CreateDateTime.HasValue)
                {
                    return this.SummaryInformation.CreateDateTime.Value.ToUniversalTime();
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
        public override IDictionary<string, object> CustomProperties
        {
            get
            {
                // Ensure file is open.
                if (!this.IsOpen)
                {
                    throw new InvalidOperationException("File is not open.");
                }

                if (this.DocumentSummaryInformation.CustomProperties == null)
                {
                    return new Dictionary<string, object>();
                }

                // Iterate manually and specify types because of dynamic COM object.
                var customProperties = new Dictionary<string, object>();
                foreach (DictionaryEntry item in this.DocumentSummaryInformation.CustomProperties)
                {
                    customProperties.Add(item.Key.ToString(), item.Value);
                }

                return customProperties;
            }
        }

        /// <summary>
        /// HPSF file properties
        /// </summary>
        public HPSFPropertiesExtractor FileProperties { get; set; }

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
                    case ".xlm":
                    case ".xlt":
                        return FileTypeEnum.MicrosoftExcel;

                    case ".ppt":
                    case ".pot":
                        return FileTypeEnum.MicrosoftPowerPoint;

                    case ".doc":
                    case ".dot":
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
                if (this.SummaryInformation.LastSaveDateTime.HasValue)
                {
                    return this.SummaryInformation.LastSaveDateTime.Value.ToUniversalTime();
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

                return this.SummaryInformation.Title;
            }
        }

        /// <summary>
        /// HPSF document summary information
        /// </summary>
        private DocumentSummaryInformation DocumentSummaryInformation { get; set; }

        /// <summary>
        /// HPSF summary information
        /// </summary>
        private SummaryInformation SummaryInformation { get; set; }

        /// <summary>
        /// Indicator if the file is writable.
        /// </summary>
        public override bool IsWritable
        {
            get { return (this._fileStream != null && this._fileStream.CanWrite); }
        }

        /// <summary>
        /// Indicator if the file is readable.
        /// </summary>
        public override bool IsReadable
        {
            get { return (this._fileStream != null && this._fileStream.CanRead); }
        }

        /// <summary>
        /// Indicator if the file is open.
        /// </summary>
        public override bool IsOpen
        {
            get { return (this._fileStream != null); }
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Closes file.
        /// </summary>
        /// <param name="saveChanges"></param>
        public override void CloseFile(bool saveChanges = false)
        {
            // Clear properties.
            this.SummaryInformation = null;
            this.DocumentSummaryInformation = null;

            // Clear file object.
            this.File = null;

            // Close stream.
            this._fileStream.Close();
        }

        /// <summary>
        /// Opens file.
        /// </summary>
        /// <param name="writable"></param>
        public override void OpenFile(bool writable = false)
        {
            // Open file stream.
            this._fileStream = new FileStream(this.Filename, FileMode.Open, FileAccess.Read);

            // Open file system.
            var fs = new POIFSFileSystem(this._fileStream);

            // Open file.
            this.File = new NpoiPropertiesOnlyDocument(fs);

            // Access properties.
            this.SummaryInformation = this.File.SummaryInformation;
            this.DocumentSummaryInformation = this.File.DocumentSummaryInformation;
        }

        #endregion Methods

    }
}
using System;
using System.Collections.Generic;
using System.IO;
using OfficeFileProperties.FileAccessors;
using OfficeFileProperties.FileAccessors.Dao;
using OfficeFileProperties.FileAccessors.Generic;
using OfficeFileProperties.FileAccessors.Npoi;
using OfficeFileProperties.FileAccessors.OpenXml;

namespace OfficeFileProperties
{
    public class OfficeFile : IFile, IDisposable
    {
        #region Constructors

        /// <summary>
        /// Accesses Microsoft Office files. Falls back to using generic filetype processing if type-specific access fails.
        /// </summary>
        /// <param name="filename">Filename to open</param>
        public OfficeFile(string filename) : this(filename, true) { }

        /// <summary>
        /// Accesses Microsoft Office files.
        /// </summary>
        /// <param name="filename">Filename to open</param>
        /// <param name="fallbackOnError">If true, errors with exceptions will be thrown rather than defaulting to basic file properties</param>
        public OfficeFile(string filename, bool fallbackOnError)
        {
            // Ensure file exists.
            if (!File.Exists(filename))
            {
                throw new FileNotFoundException($"File {filename} does not exist.");
            }

            // Get file info for new file.
            var fileInfo = new FileInfo(filename);

            // Attempt to instantiate file accessors.
            try
            {
                // Switch depending on file extension.
                switch (fileInfo.Extension.ToLower())
                {
                    case ".accdb":
                    case ".mdb":
                        // Use Dao.
                        this._fileAccessor = new DaoFile(filename);
                        break;

                    case ".doc":
                    case ".dot":
                    case ".ppt":
                    case ".pot":
                    case ".xls":
                    case ".xlm":
                    case ".xlt":
                        // Use Npoi.
                        this._fileAccessor = new NpoiFile(filename);
                        break;

                    case ".docx":
                    case ".docm":
                    case ".dotx":
                    case ".dotm":
                        // Use Docx.
                        this._fileAccessor = new DocxFile(filename);
                        break;

                    case ".pptx":
                    case ".pptm":
                    case ".potx":
                    case ".potm":
                        // Use Pptx.
                        this._fileAccessor = new PptxFile(filename);
                        break;

                    case ".xlsx":
                    case ".xlsm":
                    case ".xlst":
                        // Use Xlsx.
                        this._fileAccessor = new XlsxFile(filename);
                        break;

                    default:
                        // Use generic.
                        this._fileAccessor = new GenericFile(filename);
                        break;
                }
            }
            catch (Exception ex)
            {
                // If fallback is disabled, throw exception.
                if (!fallbackOnError)
                {
                    throw ex;
                }

                // Try using generic.
                try
                {
                    this._fileAccessor = new GenericFile(filename);
                }
                catch
                {
                    throw new Exception($"Cannot get properties from file {filename}.");
                }
            }
        }

        #endregion Constructors

        #region Fields

        /// <summary>
        /// Stores file accessor used with this object.
        /// </summary>
        private readonly IFileBase _fileAccessor;

        /// <summary>
        /// Determine if disposal has already occurred.
        /// </summary>
        private bool _disposed;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Author name
        /// </summary>
        public string Author
        {
            get => this.FileAccessor.Author;

            set => this.FileAccessor.Author = value;
        }

        /// <summary>
        /// Company name
        /// </summary>
        public string Company
        {
            get => this.FileAccessor.Company;

            set => this.FileAccessor.Company = value;
        }

        /// <summary>
        /// Comments (description)
        /// </summary>
        public string Comments
        {
            get => this.FileAccessor.Comments;

            set => this.FileAccessor.Comments = value;
        }

        /// <summary>
        /// Created Time in local time
        /// </summary>
        public DateTime? CreatedTimeLocal
        {
            get => this.FileAccessor.CreatedTimeLocal;

            set => this.FileAccessor.CreatedTimeLocal = value;
        }

        /// <summary>
        /// Created Time in UTC time
        /// </summary>
        public DateTime? CreatedTimeUtc
        {
            get => this.FileAccessor.CreatedTimeUtc;

            set => this.FileAccessor.CreatedTimeUtc = value;
        }

        /// <summary>
        /// Custom Properties
        /// </summary>
        public IDictionary<string, object> CustomProperties => this.FileAccessor.CustomProperties;

        /// <summary>
        /// Serialize Custom Properties as a string.
        /// </summary>
        public string CustomPropertiesString => this.FileAccessor.CustomPropertiesString;

        /// <summary>
        /// Accessor for manipulating files
        /// </summary>
        public IFileBase FileAccessor => this._fileAccessor;

        /// <summary>
        /// Filename
        /// </summary>
        public string Filename => this.FileAccessor.Filename;

        /// <summary>
        /// Type of file
        /// </summary>
        public FileTypeEnum FileType => this.FileAccessor.FileType;

        /// <summary>
        /// Indicator if the file is currently open
        /// </summary>
        public bool IsOpen => this.FileAccessor.IsOpen;

        /// <summary>
        /// Modified Time in local time
        /// </summary>
        public DateTime? ModifiedTimeLocal
        {
            get => this.FileAccessor.ModifiedTimeLocal;

            set => this.FileAccessor.ModifiedTimeLocal = value;
        }

        /// <summary>
        /// Modified Time in UTC time
        /// </summary>
        public DateTime? ModifiedTimeUtc
        {
            get => this.FileAccessor.ModifiedTimeUtc;

            set => this.FileAccessor.ModifiedTimeUtc = value;
        }

        /// <summary>
        /// Title
        /// </summary>
        public string Title
        {
            get => this.FileAccessor.Title;

            set => this.FileAccessor.Title = value;
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Closes file.
        /// </summary>
        public void CloseFile()
        {
            this.FileAccessor.CloseFile();
        }

        /// <summary>
        /// Dispose of object
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
        }

        /// <summary>
        /// Gets FileProperties object loaded with properties for current file.
        /// </summary>
        /// <returns>Loaded FileProperties object</returns>
        public FileProperties GetFileProperties()
        {
            return this.FileAccessor.GetFileProperties();
        }

        /// <summary>
        /// Opens file.
        /// </summary>
        /// <param name="writable">Open file in writable mode</param>
        public void OpenFile(bool writable = false)
        {
            this.FileAccessor.OpenFile(writable);
        }

        /// <summary>
        /// Dispose of object.
        /// </summary>
        /// <param name="disposing">Dispose of managed resources.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!this._disposed)
            {
                if (disposing)
                {
                    // Close file
                    this.CloseFile();
                }

                this._disposed = true;
            }
        }

        #endregion Methods
    }
}
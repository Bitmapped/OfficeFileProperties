using OfficeFileProperties.FileAccessors;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeFileProperties
{
    public class OfficeFile : IFile, IDisposable
    {

        #region Fields

        /// <summary>
        /// Stores file accessor used with this object.
        /// </summary>
        private readonly IFileBase _fileAccessor;

        /// <summary>
        /// Determine if disposal has already occurred.
        /// </summary>
        private bool _disposed = false;

        #endregion Fields

        #region Constructors

        public OfficeFile(string filename)
        {
            // Ensure file exists.
            if (!System.IO.File.Exists(filename))
            {
                throw new FileNotFoundException(String.Format("File {0} does not exist.", filename));
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
                        this._fileAccessor = new FileAccessors.Dao.DaoFile(filename);
                        break;

                    case ".doc":
                    case ".dot":
                    case ".ppt":
                    case ".pot":
                    case ".xls":
                    case ".xlm":
                    case ".xlt":
                        // Use Npoi.
                        this._fileAccessor = new FileAccessors.Npoi.NpoiFile(filename);
                        break;

                    case ".docx":
                    case ".docm":
                    case ".dotx":
                    case ".dotm":
                        // Use Docx.
                        this._fileAccessor = new FileAccessors.OpenXml.DocxFile(filename);
                        break;

                    case ".pptx":
                    case ".pptm":
                    case ".potx":
                    case ".potm":
                        // Use Pptx.
                        this._fileAccessor = new FileAccessors.OpenXml.PptxFile(filename);
                        break;

                    case ".xlsx":
                    case ".xlsm":
                    case ".xlst":
                        // Use Xlsx.
                        this._fileAccessor = new FileAccessors.OpenXml.XlsxFile(filename);
                        break;

                    default:
                        // Use generic.
                        this._fileAccessor = new FileAccessors.Generic.GenericFile(filename);
                        break;

                }
            }
            catch
            {
                // Try using generic.
                try
                {
                    this._fileAccessor = new FileAccessors.Generic.GenericFile(filename);
                }
                catch
                {
                    throw new Exception(String.Format("Cannot get properties from file {0}.", filename));
                }
            }
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// Author name
        /// </summary>
        public string Author
        {
            get
            {
                return FileAccessor.Author;
            }

            set
            {
                FileAccessor.Author = value;
            }
        }

        /// <summary>
        /// Company name
        /// </summary>
        public string Company
        {
            get
            {
                return FileAccessor.Company;
            }

            set
            {
                FileAccessor.Company = value;
            }
        }

        /// <summary>
        /// Created Time in local time
        /// </summary>
        public DateTime? CreatedTimeLocal
        {
            get
            {
                return FileAccessor.CreatedTimeLocal;
            }

            set
            {
                FileAccessor.CreatedTimeLocal = value;
            }
        }

        /// <summary>
        /// Created Time in UTC time
        /// </summary>
        public DateTime? CreatedTimeUtc
        {
            get
            {
                return FileAccessor.CreatedTimeUtc;
            }

            set
            {
                FileAccessor.CreatedTimeUtc = value;
            }
        }

        /// <summary>
        /// Custom Properties
        /// </summary>
        public IDictionary<string, object> CustomProperties
        {
            get
            {
                return FileAccessor.CustomProperties;
            }
        }

        /// <summary>
        /// Serialize Custom Properties as a string.
        /// </summary>
        public string CustomPropertiesString
        {
            get
            {
                return FileAccessor.CustomPropertiesString;
            }
        }

        /// <summary>
        /// Accessor for manipulating files
        /// </summary>
        public IFileBase FileAccessor
        {
            get
            {
                return this._fileAccessor;
            }
        }

        /// <summary>
        /// Filename
        /// </summary>
        public string Filename
        {
            get
            {
                return FileAccessor.Filename;
            }
        }

        /// <summary>
        /// Type of file
        /// </summary>
        public FileTypeEnum FileType
        {
            get
            {
                return FileAccessor.FileType;
            }
        }

        /// <summary>
        /// Indicator if the file is currently open
        /// </summary>
        public bool IsOpen
        {
            get
            {
                return FileAccessor.IsOpen;
            }
        }

        /// <summary>
        /// Modified Time in local time
        /// </summary>
        public DateTime? ModifiedTimeLocal
        {
            get
            {
                return FileAccessor.ModifiedTimeLocal;
            }

            set
            {
                FileAccessor.ModifiedTimeLocal = value;
            }
        }

        /// <summary>
        /// Modified Time in UTC time
        /// </summary>
        public DateTime? ModifiedTimeUtc
        {
            get
            {
                return FileAccessor.ModifiedTimeUtc;
            }

            set
            {
                FileAccessor.ModifiedTimeUtc = value;
            }
        }

        /// <summary>
        /// Title
        /// </summary>
        public string Title
        {
            get
            {
                return FileAccessor.Title;
            }

            set
            {
                FileAccessor.Title = value;
            }
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Closes file.
        /// </summary>
        public void CloseFile()
        {
            FileAccessor.CloseFile();
        }

        /// <summary>
        /// Dispose of object
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
        }

        /// <summary>
        /// Gets FileProperties object loaded with properties for current file.
        /// </summary>
        /// <returns>Loaded FileProperties object</returns>
        public FileProperties GetFileProperties()
        {
            return FileAccessor.GetFileProperties();
        }

        /// <summary>
        /// Opens file.
        /// </summary>
        public void OpenFile()
        {
            FileAccessor.OpenFile();
        }

        /// <summary>
        /// Dispose of object.
        /// </summary>
        /// <param name="disposing">Dispose of managed resources.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Close file
                    CloseFile();
                }

                _disposed = true;
            }
        }

        #endregion Methods
    }
}

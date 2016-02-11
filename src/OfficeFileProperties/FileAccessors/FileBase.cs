using OfficeFileProperties.Support;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeFileProperties.FileAccessors
{
    /// <summary>
    /// Abstract class for directly manipulating files.
    /// </summary>
    /// <typeparam name="T">Type of object for accessing files.</typeparam>
    public abstract class FileBase<T> : IFileBase, IDisposable
    {
        #region Fields

        /// <summary>
        /// Name of file
        /// </summary>
        private readonly string _filename;

        /// <summary>
        /// Determine if disposal has already occurred.
        /// </summary>
        private bool _disposed = false;

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

        /// <summary>
        /// Dispose of object
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
        }

        #endregion Fields

        #region Constructors

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="filename">Filename to open.</param>
        public FileBase(string filename)
        {
            // Ensure file exists.
            if (!System.IO.File.Exists(filename))
            {
                throw new FileNotFoundException(String.Format("File {0} does not exist.", filename));
            }

            // Store filename.
            this._filename = filename;
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// Author name
        /// </summary>
        public virtual string Author
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        /// <summary>
        /// Company name
        /// </summary>
        public virtual string Company
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        /// <summary>
        /// Created Time in local time
        /// </summary>
        public virtual DateTime? CreatedTimeLocal
        {
            get
            {
                if (CreatedTimeUtc.HasValue)
                {
                    return CreatedTimeUtc.Value.ToLocalTime();
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if (value.HasValue)
                {
                    CreatedTimeUtc = value.Value.ToUniversalTime();
                }
                else
                {
                    CreatedTimeUtc = null;
                }
            }
        }

        /// <summary>
        /// Created Time in UTC time
        /// </summary>
        public virtual DateTime? CreatedTimeUtc
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }


        /// <summary>
        /// Custom Properties
        /// </summary>
        public virtual IDictionary<string, string> CustomProperties
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        /// <summary>
        /// Serialize Custom Properties as a string.
        /// </summary>
        public virtual string CustomPropertiesString
        {
            get
            {
                return CustomProperties.Serialize();
            }
        }

        /// <summary>
        /// Accessor for underlying file object
        /// </summary>
        public T File { get; protected set; }

        /// <summary>
        /// Indicator if the file is currently open. Value used for internal tracking.
        /// </summary>
        private bool _isOpen = false;

        /// <summary>
        /// Filename
        /// </summary>
        public string Filename
        {
            get
            {
                return this._filename;
            }
        }

        /// <summary>
        /// Type of file
        /// </summary>
        public abstract FileTypeEnum FileType { get; }

        /// <summary>
        /// Indicator if the file is currently open
        /// </summary>
        public bool IsOpen
        {
            get
            {
                // Ensure _isOpen is true and file accessor is not null.
                return (this._isOpen && (this.File != null));
            }
            protected set
            {
                this._isOpen = value;
            }
        }

        /// <summary>
        /// Modified Time in local time
        /// </summary>
        public virtual DateTime? ModifiedTimeLocal
        {
            get
            {
                if (ModifiedTimeUtc.HasValue)
                {
                    return ModifiedTimeUtc.Value.ToLocalTime();
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if (value.HasValue)
                {
                    ModifiedTimeUtc = value.Value.ToUniversalTime();
                }
                else
                {
                    ModifiedTimeUtc = null;
                }
            }
        }

        /// <summary>
        /// Modified Time in UTC time
        /// </summary>
        public virtual DateTime? ModifiedTimeUtc
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        /// <summary>
        /// Title
        /// </summary>
        public virtual string Title
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Closes file.
        /// </summary>
        public abstract void CloseFile();

        /// <summary>
        /// Gets FileProperties object loaded with properties for current file.
        /// </summary>
        /// <returns>Loaded FileProperties object</returns>
        public virtual FileProperties GetFileProperties()
        {
            // Get new file properties object.
            var properties = new FileProperties();

            // Open file.
            OpenFile();

            // Store filename and file type.
            properties.Filename = this.Filename;
            properties.FileType = this.FileType;

            // Store properties.
            properties.Author = this.Author;
            properties.Company = this.Company;
            properties.CreatedTimeUtc = this.CreatedTimeUtc;
            properties.CustomProperties = this.CustomProperties;
            properties.ModifiedTimeUtc = this.ModifiedTimeUtc;
            properties.Title = this.Title;

            // Close file.
            CloseFile();

            return properties;
        }

        /// <summary>
        /// Opens file.
        /// </summary>
        public abstract void OpenFile();

        #endregion Methods

    }
}

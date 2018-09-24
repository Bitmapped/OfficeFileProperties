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
                    this.CloseFile();
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
        /// Comments (description)
        /// </summary>
        public virtual string Comments
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
        public virtual IDictionary<string, object> CustomProperties
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
        /// Indicator if the file has been modified.
        /// </summary>
        private bool _isDirty = false;

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
        public abstract bool IsOpen { get; }

        /// <summary>
        /// Indicator if the file is readable.
        /// </summary>
        abstract public bool IsReadable { get; }

        /// <summary>
        /// Indicator if the file is writable.
        /// </summary>
        abstract public bool IsWritable { get; }

        /// <summary>
        /// Determines if file has been modified.
        /// </summary>
        public bool IsDirty
        {
            get => (this.IsWritable && this._isDirty);
            protected set
            {
                // Check to see if we can set this.
                if (this.IsWritable)
                {
                    this._isDirty = value;
                }
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
        /// <param name="saveChanges">Save changes to this file.</param>
        public abstract void CloseFile(bool saveChanges = false);

        /// <summary>
        /// Gets FileProperties object loaded with properties for current file.
        /// </summary>
        /// <returns>Loaded FileProperties object</returns>
        public virtual FileProperties GetFileProperties()
        {
            // Open file.
            this.OpenFile();

            // Get new file properties object.
            var properties = new FileProperties
            {
                Filename = this.Filename,
                FileType = this.FileType,
                Comments = this.Comments,
                Author = this.Author,
                Company = this.Company,
                CreatedTimeUtc = this.CreatedTimeUtc,
                CustomProperties = this.CustomProperties,
                ModifiedTimeUtc = this.ModifiedTimeUtc,
                Title = this.Title
            };

            // Close file.
            this.CloseFile();

            return properties;
        }

        /// <summary>
        /// Test to ensure file is open.
        /// </summary>
        public void TestFileOpen()
        {
            // Throw exception if file is not open.
            if (!this.IsOpen)
            {
                throw new InvalidOperationException("File is not open.");
            }
        }

        public void TestFileWritable()
        {
            // Test to ensure file is open.
            this.TestFileOpen();

            // Throw exception if file is not writable.
            if (!this.IsWritable)
            {
                throw new InvalidOperationException("File is not writable.");
            }
        }

        /// <summary>
        /// Opens file.
        /// </summary>
        /// <param name="writable">Open file in writable mode</param>
        public abstract void OpenFile(bool writable = false);

        #endregion Methods

    }
}

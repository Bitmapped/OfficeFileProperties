using System;
using System.Collections.Generic;
using System.IO;
using OfficeFileProperties.Support;

namespace OfficeFileProperties.FileAccessors
{
    /// <summary>
    /// Abstract class for directly manipulating files.
    /// </summary>
    /// <typeparam name="T">Type of object for accessing files.</typeparam>
    public abstract class FileBase<T> : IFileBase, IDisposable
    {
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
                throw new FileNotFoundException(string.Format("File {0} does not exist.", filename));
            }

            // Store filename.
            this._filename = filename;
        }

        #endregion Constructors

        #region Fields

        /// <summary>
        /// Name of file
        /// </summary>
        private readonly string _filename;

        /// <summary>
        /// Determine if disposal has already occurred.
        /// </summary>
        private bool _disposed;

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

        /// <summary>
        /// Dispose of object
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
        }

        #endregion Fields

        #region Properties

        /// <summary>
        /// Author name
        /// </summary>
        public virtual string Author
        {
            get => throw new NotImplementedException();
            set => throw new NotImplementedException();
        }

        /// <summary>
        /// Company name
        /// </summary>
        public virtual string Company
        {
            get => throw new NotImplementedException();
            set => throw new NotImplementedException();
        }

        /// <summary>
        /// Comments (description)
        /// </summary>
        public virtual string Comments
        {
            get => throw new NotImplementedException();
            set => throw new NotImplementedException();
        }

        /// <summary>
        /// Created Time in local time
        /// </summary>
        public virtual DateTime? CreatedTimeLocal
        {
            get => this.CreatedTimeUtc?.ToLocalTime();
            set => this.CreatedTimeUtc = value?.ToUniversalTime();
        }

        /// <summary>
        /// Created Time in UTC time
        /// </summary>
        public virtual DateTime? CreatedTimeUtc
        {
            get => throw new NotImplementedException();
            set => throw new NotImplementedException();
        }


        /// <summary>
        /// Custom Properties
        /// </summary>
        public virtual IDictionary<string, object> CustomProperties
        {
            get => throw new NotImplementedException();
            set => throw new NotImplementedException();
        }

        /// <summary>
        /// Serialize Custom Properties as a string.
        /// </summary>
        public virtual string CustomPropertiesString => this.CustomProperties.Serialize();

        /// <summary>
        /// Accessor for underlying file object
        /// </summary>
        public T File { get; protected set; }

        /// <summary>
        /// Filename
        /// </summary>
        public string Filename => this._filename;

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
        public abstract bool IsReadable { get; }

        /// <summary>
        /// Indicator if the file is writable.
        /// </summary>
        public abstract bool IsWritable { get; }

        /// <summary>
        /// Modified Time in local time
        /// </summary>
        public virtual DateTime? ModifiedTimeLocal
        {
            get => this.ModifiedTimeUtc?.ToLocalTime();
            set => this.ModifiedTimeUtc = value?.ToUniversalTime();
        }

        /// <summary>
        /// Modified Time in UTC time
        /// </summary>
        public virtual DateTime? ModifiedTimeUtc
        {
            get => throw new NotImplementedException();
            set => throw new NotImplementedException();
        }

        /// <summary>
        /// Title
        /// </summary>
        public virtual string Title
        {
            get => throw new NotImplementedException();
            set => throw new NotImplementedException();
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

        /// <summary>
        /// Test to ensure file is writable.
        /// </summary>
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
        /// Test to ensure file is readable.
        /// </summary>
        public void TestFileRedable()
        {
            // Test to ensure file is open.
            this.TestFileOpen();

            // Throw exception if file is not writable.
            if (!this.IsReadable)
            {
                throw new InvalidOperationException("File is not readable.");
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
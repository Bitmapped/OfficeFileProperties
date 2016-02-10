using OfficeFileProperties.Support;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeFileProperties.FileAccessors
{
    public abstract class FileBase<T>
    {
        #region Fields

        /// <summary>
        /// Name of file
        /// </summary>
        private readonly string _filename;

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
        /// Created Date in local time
        /// </summary>
        public virtual DateTime? CreatedDateLocal
        {
            get
            {
                if (CreatedDateUtc.HasValue)
                {
                    return CreatedDateUtc.Value.ToLocalTime();
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
                    CreatedDateUtc = value.Value.ToUniversalTime();
                }
                else
                {
                    CreatedDateUtc = null;
                }
            }
        }

        /// <summary>
        /// Created date in UTC time
        /// </summary>
        public virtual DateTime? CreatedDateUtc
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
        public virtual Dictionary<string, string> CustomProperties
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
        public T FileAccessor { get; protected set; }

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
        public bool IsOpen { get; protected set; }

        /// <summary>
        /// Modified Date in local time
        /// </summary>
        public virtual DateTime? ModifiedDateLocal
        {
            get
            {
                if (ModifiedDateUtc.HasValue)
                {
                    return ModifiedDateUtc.Value.ToLocalTime();
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
                    ModifiedDateUtc = value.Value.ToUniversalTime();
                }
                else
                {
                    ModifiedDateUtc = null;
                }                
            }
        }

        /// <summary>
        /// Modified Date in UTC time
        /// </summary>
        public virtual DateTime? ModifiedDateUtc
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
            properties.CreatedDateUtc = this.CreatedDateUtc;
            properties.CustomProperties = this.CustomProperties;
            properties.ModifiedDateUtc = this.ModifiedDateUtc;
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

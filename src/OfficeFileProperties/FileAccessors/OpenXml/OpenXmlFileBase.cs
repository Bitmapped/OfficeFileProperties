using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeFileProperties.FileAccessors.OpenXml
{
    public abstract class OpenXmlFileBase<T> : FileBase<T> where T : OpenXmlPackage
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="filename">Filename to open.</param>
        public OpenXmlFileBase(string filename) : base(filename)
        { }

        /// <summary>
        /// Author name
        /// </summary>
        public override string Author
        {
            get
            {
                // Ensure file is open.
                this.TestFileOpen();

                return this.File.PackageProperties.Creator;
            }
            set
            {
                // Ensure file is writable.
                this.TestFileWritable();

                // Set author.
                this.File.PackageProperties.Creator = value;
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
                this.TestFileOpen();

                return this.File.PackageProperties.Description;
            }
            set
            {
                // Ensure file is writable.
                this.TestFileWritable();

                // Set comments.
                this.File.PackageProperties.Description = value;
            }
        }

        /// <summary>
        /// Created date in UTC time
        /// </summary>
        public override DateTime? CreatedTimeUtc
        {
            get
            {
                // Ensure file is open.
                this.TestFileOpen();

                return this.File.PackageProperties.Created?.ToUniversalTime();
            }
            set
            {
                // Ensure file is writable.
                this.TestFileWritable();

                // Set created time.
                this.File.PackageProperties.Created = value;
            }
        }

        /// <summary>
        /// Determine if file is open.
        /// </summary>
        public override bool IsOpen
        {
            get
            {
                switch (this.File?.FileOpenAccess)
                {
                    case FileAccess.ReadWrite:
                    case FileAccess.Write:
                    case FileAccess.Read:
                        return true;

                    default:
                        return false;
                }
            }
        }

        /// <summary>
        /// Indicator if the file is readable.
        /// </summary>
        public override bool IsReadable
        {
            get
            {
                switch (this.File?.FileOpenAccess)
                {
                    case FileAccess.Read:
                    case FileAccess.ReadWrite:
                        return true;

                    default:
                        return false;
                }
            }
        }


        /// <summary>
        /// Indicator if the file is writable.
        /// </summary>
        public override bool IsWritable
        {
            get
            {
                switch (this.File?.FileOpenAccess)
                {
                    case FileAccess.ReadWrite:
                    case FileAccess.Write:
                        return true;

                    default:
                        return false;
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
                // Ensure file is open.
                this.TestFileOpen();

                return this.File.PackageProperties.Modified?.ToUniversalTime();
            }
            set
            {
                // Ensure file is writable.
                this.TestFileWritable();

                // Set modified time.
                this.File.PackageProperties.Modified = value;
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
                this.TestFileOpen();

                return this.File.PackageProperties.Title;
            }
            set
            {
                // Ensure file is writable.
                this.TestFileWritable();

                // Set title.
                this.File.PackageProperties.Title = value;
            }
        }

        /// <summary>
        /// Closes file.
        /// </summary>
        public override void CloseFile()
        {
            // Close file if it still is accessible.
            if (this.File != null)
            {
                // Close file.
                this.File.Close();

                // Dispose of file.
                this.File.Dispose();
            }

            // Clear file object.
            this.File = null;
        }
    }
}
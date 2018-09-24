using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using NPOI.SS.Formula.Functions;

namespace OfficeFileProperties.FileAccessors.OpenXml
{
    abstract public class OpenXmlFileBase<T> : FileBase<T> where T : OpenXmlPackage
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="filename">Filename to open.</param>
        public OpenXmlFileBase(string filename) : base(filename)
        { }

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

        public override bool IsWritable
        {
            get
            {
                // Check if file is open.
                if (!this.IsOpen)
                {
                    return false;
                }

                // See if file has been opened in a mode that allows writing.
                if ((this.File.FileOpenAccess == FileAccess.ReadWrite) || (this.File.FileOpenAccess == FileAccess.Write))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        /// <summary>
        /// Closes file.
        /// </summary>
        /// <param name="saveChanges"></param>
        public override void CloseFile(bool saveChanges = false)
        {
            // If file has changes, is writable, and is to be saved, save it.
            if (saveChanges && this.IsWritable && this.IsDirty)
            {
                this.File.Save();

                // Mark file as not dirty.
                this.IsDirty = false;
            }

            // Mark file as closed.
            this.IsOpen = false;

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

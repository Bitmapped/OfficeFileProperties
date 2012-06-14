using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace OfficeFileProperties.File.Generic
{
    /// <summary>
    /// Obtain properties of generic files using basic Windows methods.
    /// </summary>
    class GenericFile : IFile
    {
        // Define private variables.
        private string filename;
        private FileInfo file;
        private GenericFileProperties fileProperties;

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
        public GenericFile()
        {
            ClearProperties();
        }

        /// <summary>
        /// Constructor, also loads specified file.
        /// </summary>
        /// <param name="filename">File to load.</param>
        public GenericFile(string filename) : this()
        {
            // Load specified file.
            this.LoadFile(filename);
        }

        /// <summary>
        /// Forget about loaded file.
        /// </summary>
        private void CloseFile()
        {
            // Clear fileInfo object.
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
            this.file = new FileInfo(filename);

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
            this.fileProperties = new GenericFileProperties();

            // filename
            this.fileProperties.filename = this.filename;
            
            // createdTimeUtc
            this.fileProperties.createdTimeUtc = this.file.CreationTimeUtc;

            // modifiedTimeUtc
            this.fileProperties.modifiedTimeUtc = this.file.LastWriteTimeUtc;

            // Mark properties as loaded.
            this.fileProperties.fileLoaded = true;
        }

    }
}
